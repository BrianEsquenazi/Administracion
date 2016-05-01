VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsFicTer 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Ficha de Stock de Producto Terminado"
   ClientHeight    =   8280
   ClientLeft      =   180
   ClientTop       =   405
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   8280
   ScaleWidth      =   11655
   Begin VB.ComboBox Tipo 
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
      TabIndex        =   57
      Top             =   1680
      Width           =   2415
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
      Index           =   10
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   4680
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
      Index           =   9
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   4320
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   4080
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4080
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4080
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4080
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
      Index           =   5
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   4080
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4080
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4080
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
      Index           =   8
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4080
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WMuestra 
      Height          =   4815
      Left            =   120
      TabIndex        =   37
      Top             =   3240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16777088
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   0
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wloteter.rpt"
   End
   Begin VB.Frame StockCons 
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
      Left            =   3240
      TabIndex        =   26
      Top             =   2040
      Width           =   8055
      Begin VB.Label WStock7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   6480
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Stock7 
         Caption         =   "Stock"
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
         Left            =   5400
         TabIndex        =   55
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WStock6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   3600
         TabIndex        =   54
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Stock6 
         Caption         =   "Stock"
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
         Left            =   2520
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Stock4 
         Caption         =   "Stock"
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
         Left            =   2520
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Stock5 
         Caption         =   "Stock"
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
         Left            =   2520
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label WStock4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   3600
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WStock5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   3600
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label WStock3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   1200
         TabIndex        =   32
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label WStock2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   1200
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label WStock1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Stock3 
         Caption         =   "Stock"
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
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Stock2 
         Caption         =   "Stock"
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
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Stock1 
         Caption         =   "Stock"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox RE 
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Nk 
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox StkProceso 
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   " "
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Deposito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   " "
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Unidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   " "
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox XStock 
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   " "
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox XSalidas 
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox XEntradas 
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox XInicial 
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " "
      Top             =   120
      Width           =   1575
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327680
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-#####-###"
      PromptChar      =   " "
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   840
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
      ItemData        =   "consficter.frx":0000
      Left            =   360
      List            =   "consficter.frx":0007
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   3480
      TabIndex        =   2
      Top             =   120
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton AvisoError 
      Caption         =   "Sistema sin Conexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      Picture         =   "consficter.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport listado2 
      Left            =   -120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "listado2.rpt"
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   51
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Faltante"
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
      Left            =   240
      TabIndex        =   50
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Pedidos Pendientes"
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
      Left            =   240
      TabIndex        =   49
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   48
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Stock Re"
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
      TabIndex        =   24
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Stock NK"
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
      TabIndex        =   19
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Stock en Proceso"
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
      TabIndex        =   18
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Deposito"
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
      TabIndex        =   17
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Unidad de Medida"
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
      TabIndex        =   16
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo Final"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Salidas"
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
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Entradas"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Saldo Inicial"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   120
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
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsFicTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Clave As String
Private Vector(2000, 12) As String
Dim Termi As String

Dim rsttotalpt As Recordset
Dim sptotalpt As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstEntdev As Recordset
Dim spEntdev As String

Dim XParam As String
Dim Auxiliar(10000, 6) As String
Dim XLote(12, 2) As String
Dim WXEntrada As Double
Dim WXSalida As Double
Dim WXInicial As Double
Dim WXStock As Double
Dim WSaldo As Double

Dim ZMes As String
Dim ZAno As String

Dim ZZZZCanti1 As Double
Dim ZZZZCanti2 As Double
Dim ZZZZCanti3 As Double
Dim ZZZZCanti4 As Double
Dim ZZZZCanti5 As Double

Private Sub cmdClose_Click()
    Terminado.SetFocus
    PrgConsFicTer.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
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
            
    Pantalla.Visible = True

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaTer
End Sub

Private Sub Label11_Click()

    listado2.WindowTitle = "Listado de Ordenes de Compra por Proveedor"
    listado2.WindowTop = 0
    listado2.WindowLeft = 0
    listado2.WindowWidth = Screen.Width
    listado2.WindowHeight = Screen.Height
  
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    listado2.SQLQuery = "SELECT Pedido.Clave, Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Cliente.Razon, Pedido.Cantidad " _
                       + "From " _
                       + DSQ + ".dbo.Pedido Pedido, " _
                       + DSQ + ".dbo.Cliente Cliente " _
                       + "Where " _
                       + "Pedido.Terminado = '" + Terminado.Text + "' AND " _
                       + "Pedido.Cliente = Cliente.Cliente AND " _
                       + "Pedido.Facturado < Pedido.Cantidad "
   
    listado2.Connect = Connect()
    listado2.Action = 1
    
    XEmpresa = Wempresa
    Call Conecta_Empresa

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Terminado.Text = rstTerminado!Codigo
        DesTerminado.Caption = rstTerminado!Descripcion
        Unidad.Text = rstTerminado!Unidad
        Deposito.Text = rstTerminado!Deposito
        StkProceso.Text = Pusing("###,###,###.##", Str$(rstTerminado!Proceso))
        rstTerminado.Close
        Label11.Caption = "0"
        Label11.Caption = Pusing("###,###,###.##", Label11.Caption)
        Label15.Caption = "0"
        Label15.Caption = Pusing("###,###,###.##", Label15.Caption)
        
        Call Proceso_Click
            Else
        Terminado.Text = Claveven$
    End If
    Rem Terminado.SetFocus
    
End Sub

Private Sub StkProceso_DBLCLICK()
    
    
    Sql1 = "UPDATE Hoja SET "
    Sql2 = " Realant = 0"
    Sql3 = " Where Realant IS NULL"
    spHoja = Sql1 + Sql2 + Sql3
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Listado.WindowTitle = "Listado de Hoja de Produccion Pendirentes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Hoja.Marca} <> " + Chr$(34) + "X" + Chr$(34) + " and {Hoja.Real} = 0 and {Hoja.Teorico} <> 0 and {Hoja.Renglon} = 1 and {Hoja.Producto} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Renglon, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, Hoja.Real, Hoja.FechaIngOrd, Hoja.Marca, Hoja.Realant,  " _
                        + "Terminado.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.Hoja Hoja, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Hoja.Producto = Terminado.Codigo AND " _
                        + "Hoja.Renglon = 1 AND " _
                        + "Hoja.Producto >= '" + Terminado.Text + "' AND " _
                        + "Hoja.Producto <= '" + Terminado.Text + "' AND " _
                        + "Hoja.Teorico <> 0 AND " _
                        + "Hoja.Real = 0 AND " _
                        + "Hoja.Realant = 0 AND " _
                        + "Hoja.Marca <> 'X'"
                        

    Listado.DataFiles(2) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    WListado = Listado.ReportFileName
    Listado.ReportFileName = "Wlisthojapend.rpt"
    Listado.Action = 1
    Listado.ReportFileName = WListado

End Sub

Private Sub Tipo_Click()
    If Terminado.Text <> "  -     -   " Then
        Call Proceso_Click
    End If
End Sub

Private Sub WMuestra_DblClick()

    On Error GoTo WError
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
        rstTerminado.Close
    End If

    WMuestra.Col = 8
    Pasalote = WMuestra.Text
    
    Erase Vector
    Renglon = 0

    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
            
Rem DADA
Rem DADA
Rem DADA
            
    XParam = "'" + Pasalote + "'"
    spHoja = "ListaHoja" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        WProducto = rstHoja!Producto
        WCantidad = rstHoja!Real
        WCantidadII = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
        If WCantidadII <> 0 Then
            WCantidad = WCantidadII
        End If
        Rem WCantidad = WCantidad + WCantidadII
        WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
        WFechaFinal = Trim(WFechaFinal)
        If WFechaFinal <> "" Then
            WFecha = WFechaFinal
                Else
            WFecha = rstHoja!Fecha
        End If
        Rem WFecha = rstHoja!Fecha
        WHoja = rstHoja!Hoja
        WSaldo = rstHoja!Saldo
                
        With rstFichaTer
                
            .AddNew
            !Terminado = WProducto
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WHoja
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = 0
            !Observaciones = ""
            !Lista1 = "Hoja"
            !Lista2 = ""
            !Lote = WHoja
            !Saldo = WSaldo
            .Update
        End With
        
        rstHoja.Close
        
            Else
            
        XParam = "'" + Pasalote + "'"
        spMovguia = "ListaMovguiaLoteSolo" + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            WProducto = rstMovguia!Terminado
            WProducto = Terminado.Text
            rstMovguia.Close
        End If
        
    End If
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
    spEstadistica = "ListaEstadisticaDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstEstadistica!Marca = "X" Then
                If rstEstadistica!Marca = "X" Then
                
                        Else
                        
                    Erase XLote
    
                    ZZZZCanti1 = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    ZZZZCanti2 = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                    ZZZZCanti3 = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                    ZZZZCanti4 = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                    ZZZZCanti5 = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                    
                    XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                    XLote(1, 2) = Str$(ZZZZCanti1)
                    XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                    XLote(2, 2) = Str$(ZZZZCanti2)
                    XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                    XLote(3, 2) = Str$(ZZZZCanti3)
                    XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                    XLote(4, 2) = Str$(ZZZZCanti4)
                    XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                    XLote(5, 2) = Str$(ZZZZCanti5)
                    
                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                    
                    If Len(Trim(WLoteAdicional)) = 98 Then
                        XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                        XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                        XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                        XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                        XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                        XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                        XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                        XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                        XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                        XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                        XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                        XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                        XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                        XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                            Else
                        XLote(6, 1) = ""
                        XLote(6, 2) = "0"
                        XLote(7, 1) = ""
                        XLote(7, 2) = "0"
                        XLote(8, 1) = ""
                        XLote(8, 2) = "0"
                        XLote(9, 1) = ""
                        XLote(9, 2) = "0"
                        XLote(10, 1) = ""
                        XLote(10, 2) = "0"
                        XLote(11, 1) = ""
                        XLote(11, 2) = "0"
                        XLote(12, 1) = ""
                        XLote(12, 2) = "0"
                            
                    End If
                        
                    If XLote(1, 2) = 0 Then
                        XLote(1, 2) = rstEstadistica!Cantidad
                    End If
                    
                    For ZZCiclo = 1 To 12
                    
                        If Val(XLote(ZZCiclo, 1)) = Val(Pasalote) Then
                        
                            WTipo = rstEstadistica!Tipo
                            WTerminado = rstEstadistica!Articulo
                            WSalida = XLote(ZZCiclo, 2)
                            WFecha = rstEstadistica!Fecha
                            WNumero = rstEstadistica!Numero
                            WImpre1 = rstEstadistica!Cliente
                            
                            Renglon = Renglon + 1
                        
                            Vector(Renglon, 1) = WTipo
                            Vector(Renglon, 2) = WTerminado
                            Vector(Renglon, 3) = WSalida
                            Vector(Renglon, 4) = WFecha
                            Vector(Renglon, 5) = WNumero
                            Vector(Renglon, 6) = WImpre1
                        
                        End If
                        
                    Next ZZCiclo
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
    For Da = 1 To Renglon
    
        WTipo = Val(Vector(Da, 1))
        WTerminado = Vector(Da, 2)
        WSalida = Val(Vector(Da, 3))
        WFecha = Vector(Da, 4)
        WNumero = Vector(Da, 5)
        WImpre1 = Vector(Da, 6)
        
        XEmpresa = Wempresa
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WImpre2 = rstCliente!Razon
            rstCliente.Close
                Else
            WImpre2 = ""
        End If
        
        Call Conecta_Empresa
                
        With rstFichaTer
                
                .AddNew
                !Terminado = WTerminado
                !Fecha = WFecha
                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Tipo = 0
                !Numero = WNumero
                !Inicial = 0
                If Val(WTipo) = 1 Then
                    !Entrada = 0
                    !Salida = WSalida
                    !Lista1 = "Factura"
                            Else
                    !Salida = 0
                    !Entrada = Abs(WSalida)
                    !Lista1 = "Devol"
                End If
                !Observaciones = ""
                !Lista2 = WImpre1 + " " + Left$(WImpre2, 23)
                !Lote = Val(Pasalote)
                !Saldo = 0
                .Update
        End With
    Next Da
    
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
    spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                
                Rem If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                If rstHoja!Marca = "X" Then
                
                    Else
                
                If rstHoja!Tipo = "T" Then
                
                    WTerminado = rstHoja!Terminado
                    WCantidad = rstHoja!Cantidad
                    WCanti1 = rstHoja!Canti1
                    WCanti2 = rstHoja!Canti2
                    WCanti3 = rstHoja!Canti3
                
                    WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
                    WFechaFinal = Trim(WFechaFinal)
                    If WFechaFinal <> "" Then
                        WFecha = WFechaFinal
                            Else
                        WFecha = rstHoja!Fecha
                    End If
                    
                    WHoja = rstHoja!Hoja
                    
                    
                    If rstHoja!lote1 = Val(Pasalote) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = WCanti1
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Pasalote
                            !Saldo = 0
                            .Update
                        End With
                        
                    End If
                    
                    If rstHoja!lote2 = Val(Pasalote) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = WCanti2
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Pasalote
                            !Saldo = 0
                            .Update
                        End With
                        
                    End If
                    
                    If rstHoja!lote3 = Val(Pasalote) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = WCanti3
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Pasalote
                            !Saldo = 0
                            .Update
                        End With
                        
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
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
    spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstMovvar!Marca = "X" Then
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                If rstMovvar!Tipo = "T" Then
                
                    WTerminado = rstMovvar!Terminado
                    WCantidad = rstMovvar!Cantidad
                    WFecha = rstMovvar!Fecha
                    WCodigo = rstMovvar!Codigo
                    WMovi = rstMovvar!Movi
                    WTipomov = Val(rstMovvar!Tipomov)
                    WObservaciones = rstMovvar!Observaciones
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    If Val(WLote) = Val(Pasalote) Then

                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            If WTipomov = 1 Or WTipomov = 2 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = 0
                            .Update
                        End With
                        
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
        rstMovvar.Close
    End If
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                Rem If rstMovguia!Marca = "X" Then
                If rstMovguia!Marca = "X" Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    WTerminado = rstMovguia!Terminado
                    WCantidadAnt = IIf(IsNull(rstMovguia!Cantidadant), "0", rstMovguia!Cantidadant)
                    WCantidad = rstMovguia!Cantidad
                    If WCantidadAnt <> 0 Then
                        WCantidad = WCantidadAnt
                            Else
                        WCantidad = WCantidad
                    End If
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    Rem WObservaciones = rstMovvar!Observaciones
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    
                    If WMovi = "S" Then
                            Select Case WDestino
                                Case 1
                                    WObservaciones = "Envio a Surfactan"
                                Case 2
                                    WObservaciones = "Envio a Pellital"
                                Case 3
                                    WObservaciones = "Envio a Surfactan II"
                                Case 4
                                    WObservaciones = "Envio a Pellital II"
                                Case 5
                                    WObservaciones = "Envio a Surfactan III"
                                Case 6
                                    WObservaciones = "Envio a Surfactan IV"
                                Case 7
                                    WObservaciones = "Envio a Surfactan V"
                                Case 8
                                    WObservaciones = "Envio a Pellital V"
                                Case 9
                                    WObservaciones = "Envio a Pellital IV"
                                Case 10
                                    WObservaciones = "Envio a Surfactan VI"
                                Case 11
                                    WObservaciones = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Partida
                            WSaldo = 0
                            
                                Else
                                
                            Select Case WTipomov
                                Case 1
                                    WObservaciones = "Recep. Surfactan"
                                Case 2
                                    WObservaciones = "Recep. Pellital"
                                Case 3
                                    WObservaciones = "Recep. Surfactan II"
                                Case 4
                                    WObservaciones = "Recep. Pellital II"
                                Case 5
                                    WObservaciones = "Recep. Surfactan III"
                                Case 6
                                    WObservaciones = "Recep. Surfactan IV"
                                Case 7
                                    WObservaciones = "Recep. Surfactan V"
                                Case 8
                                    WObservaciones = "Recep. Pellital V"
                                Case 9
                                    WObservaciones = "Recep. Pellital IV"
                                Case 10
                                    WObservaciones = "Recep. Surfactan VI"
                                Case 11
                                    WObservaciones = "Recep. Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Lote
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            
                    End If
                    
                    If WLote = Val(Pasalote) Then
                        
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            !Lista1 = "Guia In"
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = WSaldo
                            .Update
                        End With
                    
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
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
    spConsig = "ListaConsigTerminado" + XParam
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
    
        With rstConsig
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstConsig!Marca <> "X" Then
                If rstConsig!Marca <> "X" Then
                
                    WTerminado = rstConsig!Terminado
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WFecha = rstConsig!Fecha
                    WCodigo = rstConsig!Numero
                    WCliente = rstConsig!Cliente
                    WObservaciones = rstConsig!Observaciones
                    WLote = rstConsig!Lote
                    
                    If WCantidad <> 0 Then
                    
                        If WLote = Val(Pasalote) Then

                            With rstFichaTer
                                .AddNew
                                !Terminado = WTerminado
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WCodigo
                                !Inicial = 0
                                !Entrada = 0
                                !Salida = WCantidad
                                !Observaciones = WCliente
                                !Lista1 = "Rem.Con."
                                !Lista2 = Left$(WObservaciones, 30)
                                !Lote = Pasalote
                                !Saldo = 0
                                .Update
                            End With
                        
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
        rstConsig.Close
    End If
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
    spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstMovlab!Marca = "X" Then
                If rstMovlab!Marca = "X" Then
                
                        Else
                
                If rstMovlab!Tipo = "T" Then
                
                    WTerminado = rstMovlab!Terminado
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = rstMovlab!Lote
                    
                    If WLote = Val(Pasalote) Then

                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            !Lista1 = "Mov.Lab"
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = 0
                            .Update
                        End With
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
        rstMovlab.Close
    End If
    
    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WTerminado = !Terminado
                WDescripcion = ""
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                !Descripcion = WDescripcion
                
                If Left$(!Lista1, 8) = "Rem.Con." Then
                
                    XEmpresa = Wempresa
                    Select Case Val(Wempresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            Wempresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            Wempresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                
                    spCliente = "ConsultaCliente " + "'" + Left$(!Observaciones, 6) + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WLista2 = Left$(rstCliente!Razon, 30)
                        rstCliente.Close
                            Else
                        WLista2 = ""
                    End If
                    
                    Call Conecta_Empresa
                    
                    !Lista2 = WLista2
                    
                End If
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Ficha de Lote de Productos Terminados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    Listado.Destination = 0
    Listado.DataFiles(0) = Wempresa + "auxi.mdb"
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Rem Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub

WError:

    Resume Next
    
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgConsFicTer.Caption = "Consulta de Ficha de Stock de Producto Terminado :  " + !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "Con Saldo"
    Tipo.AddItem "Todos los movimientos"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Proceso_Click()
        
    If Tipo.ListIndex = 0 Then
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        XInicial.Visible = False
        XEntradas.Visible = False
        XSalidas.Visible = False
        XStock.Visible = False
            Else
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        XInicial.Visible = True
        XEntradas.Visible = True
        XSalidas.Visible = True
        XStock.Visible = True
    End If
        
    Terminado.Text = UCase(Terminado.Text)
    
    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
         Stock1.Caption = "Pta I"
         Stock2.Caption = "Pta II"
         Stock3.Caption = "Pta III"
         Stock4.Caption = "Pta IV"
         Stock5.Caption = "Pta V"
         Stock6.Caption = "Pta VI"
         Stock7.Caption = "Pta VII"
            Else
         Stock1.Caption = "Pta I"
         Stock2.Caption = "Pta II"
         Stock3.Caption = "Pta V"
         Stock4.Caption = "Pta IV"
         Stock5.Caption = ""
         Stock6.Caption = ""
         Stock7.Caption = ""
    End If
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
        
            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock1.Caption = "0"
            End If
        
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock3.Caption = "0"
            End If
            
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock4.Caption = "0"
            End If
            
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock5.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock5.Caption = "0"
            End If
            
            Wempresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock6.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock6.Caption = "0"
            End If
            
            Wempresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock7.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock7.Caption = "0"
            End If
    
            Select Case Val(XEmpresa)
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
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
    
        Case 2, 4, 8, 9
        
            XEmpresa = Wempresa
    
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                      Else
                WStock1.Caption = "0"
            End If
    
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                      Else
                WStock3.Caption = "0"
            End If
            
            
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spTerminado = "Consultaterminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                rstTerminado.Close
                      Else
                WStock4.Caption = "0"
            End If
            
            
            Select Case Val(XEmpresa)
                Case 2
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
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
                Case Else
            End Select
    
        Case Else
    End Select
    
    On Error GoTo 0
    
    WStock1.Caption = Pusing("###,###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###,###.##", WStock4.Caption)
    WStock5.Caption = Pusing("###,###,###.##", WStock5.Caption)
    WStock6.Caption = Pusing("###,###,###.##", WStock6.Caption)
    WStock7.Caption = Pusing("###,###,###.##", WStock7.Caption)
    
    WXInicial = 0
    WXEntradas = 0
    WXSalidas = 0
    WXStock = 0
    
    Call Limpia_Vector
    
    Renglon = 0
    
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
        WTerminado = rstTerminado!Codigo
        WInicial = rstTerminado!Inicial
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
                                        
        Renglon = Renglon + 1
                
        WMuestra.TextMatrix(Renglon, 1) = WFechaCierre
        WMuestra.TextMatrix(Renglon, 2) = ""
        WMuestra.TextMatrix(Renglon, 3) = ""
        WMuestra.TextMatrix(Renglon, 4) = "Saldo Inicial"
        WMuestra.TextMatrix(Renglon, 5) = Pusing("###,###,###.##", Str$(rstTerminado!Inicial))
        WMuestra.TextMatrix(Renglon, 6) = ""
        WMuestra.TextMatrix(Renglon, 7) = ""
        WMuestra.TextMatrix(Renglon, 8) = ""
        WMuestra.TextMatrix(Renglon, 9) = ""
        WMuestra.TextMatrix(Renglon, 10) = ""
                
        WXInicial = rstTerminado!Inicial
        
        rstTerminado.Close
        
    End If
                
                
    Rem dada
    Rem PROCESA LAS ESTADISTICAS
    Rem dada
    
    If Tipo.ListIndex = 1 Then
        
        Select Case Left$(Terminado.Text, 2)
            Case "PT"
                Rem PROCESA LAS ESTADISTICAS
        
                Erase Vector
                Lugar = 0
        
                Sql1 = "Select Estadistica.Marca, Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
                Sql2 = " FROM Estadistica"
                Sql3 = " Where Estadistica.Articulo >= " + "'" + Terminado.Text + "'"
                Sql4 = " and Estadistica.Articulo <= " + "'" + Terminado.Text + "'"
                Sql5 = " and Estadistica.Marca <> " + "'" + "X" + "'"
                spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
        
                    With rstEstadistica
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                
                                If .EOF = True Then
                                    Exit Do
                                End If
                    
                                WImpre2 = rstEstadistica!Cliente
                                
                                WTipo = rstEstadistica!Tipo
                                WTerminado = rstEstadistica!Articulo
                                WSalida = rstEstadistica!Cantidad
                                WFecha = rstEstadistica!Fecha
                                WNumero = rstEstadistica!Numero
                                WImpre1 = rstEstadistica!Cliente
                                WCliente = rstEstadistica!Cliente
                    
                                Erase XLote
                    
                                XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                                XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                                XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                                XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                                XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                                XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                                XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                                XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                                XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                                XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                                
                                WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                                
                                If Len(Trim(WLoteAdicional)) = 98 Then
                                    XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                                    XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                                    XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                                    XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                                    XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                                    XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                                    XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                                    XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                                    XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                                    XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                                    XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                                    XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                                    XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                                    XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                End If
                    
                                If XLote(1, 2) = 0 Then
                                    XLote(1, 2) = rstEstadistica!Cantidad
                                End If
                    
                                For x = 1 To 12
                    
                                    If Val(XLote(x, 2)) <> 0 Then
                    
                                        WSalida = XLote(x, 2)
                                        Lugar = Lugar + 1
                        
                                        Vector(Lugar, 1) = WFecha
                                        Select Case Val(WTipo)
                                            Case 1
                                                Vector(Lugar, 2) = "Fac"
                                                Vector(Lugar, 5) = ""
                                                Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WSalida))
                                                WXSalidas = WXSalidas + WSalida
                                            Case Else
                                                Vector(Lugar, 2) = "Dev"
                                                Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(Abs(WSalida)))
                                                Vector(Lugar, 6) = ""
                                                WXEntradas = WXEntradas + Abs(WSalida)
                                        End Select
                                        Vector(Lugar, 3) = WNumero
                                        Vector(Lugar, 4) = WImpre2
                                        If Left$(rstEstadistica!Remito, 1) = "C" Then
                                            Vector(Lugar, 7) = Str$(Val(Mid$(rstEstadistica!Remito, 2, 9)))
                                                Else
                                            Vector(Lugar, 7) = Str$(Val(rstEstadistica!Remito))
                                        End If
                                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        Vector(Lugar, 9) = WImpre1
                                        Vector(Lugar, 10) = Trim(Str$(Val(XLote(x, 1))))
                                        Vector(Lugar, 11) = ""
                                    End If
                    
                                Next x
                    
                                .MoveNext
                    
                                If .EOF = True Then
                                    Exit Do
                                End If
                    
                            Loop
                        End If
                
                    End With
            
                    rstEstadistica.Close
            
                End If
                
            Case "NK"
                Rem PROCESA LAS ESTADISTICAS
        
                Erase Vector
                Lugar = 0
        
                Sql1 = "Select Estadistica.Marca, Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
                Sql2 = " FROM Estadistica"
                Sql3 = " Where Estadistica.Articulo >= " + "'" + "PT" + Mid$(Terminado.Text, 3, 10) + "'"
                Sql4 = " and Estadistica.Articulo <= " + "'" + "PT" + Mid$(Terminado.Text, 3, 10) + "'"
                Sql5 = " and Estadistica.Marca <> " + "'" + "X" + "'"
                Sql6 = " and Estadistica.Tipo <> " + "'" + "1" + "'"
                spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
        
                    With rstEstadistica
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                
                                If .EOF = True Then
                                    Exit Do
                                End If
                    
                                WImpre2 = rstEstadistica!Cliente
                    
                                WTipo = rstEstadistica!Tipo
                                WTerminado = rstEstadistica!Articulo
                                WSalida = rstEstadistica!Cantidad
                                WFecha = rstEstadistica!Fecha
                                WNumero = rstEstadistica!Numero
                                WImpre1 = rstEstadistica!Cliente
                                WCliente = rstEstadistica!Cliente
                    
                                Erase XLote
                    
                                XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                                XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                                XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                                XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                                XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                                XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                                XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                                XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                                XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                                XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                                
                                WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                                
                                If Len(Trim(WLoteAdicional)) = 98 Then
                                    XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                                    XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                                    XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                                    XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                                    XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                                    XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                                    XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                                    XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                                    XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                                    XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                                    XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                                    XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                                    XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                                    XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                End If
                    
                                If XLote(1, 2) = 0 Then
                                    XLote(1, 2) = rstEstadistica!Cantidad
                                End If
                    
                                For x = 1 To 12
                    
                                    If Val(XLote(x, 2)) <> 0 Then
                    
                                        WSalida = XLote(x, 2)
                                        Lugar = Lugar + 1
                        
                                        Vector(Lugar, 1) = WFecha
                                        Vector(Lugar, 2) = "Dev"
                                        Vector(Lugar, 5) = ""
                                        Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WSalida))
                                        WXSalidas = WXSalidas + WSalida
                                        Vector(Lugar, 3) = WNumero
                                        Vector(Lugar, 4) = WImpre2
                                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        Vector(Lugar, 9) = WImpre1
                                        Vector(Lugar, 10) = Trim(Str$(Val(XLote(x, 1))))
                                        Vector(Lugar, 11) = ""
                                    End If
                    
                                Next x
                    
                                .MoveNext
                    
                                If .EOF = True Then
                                    Exit Do
                                End If
                    
                            Loop
                        End If
                
                    End With
            
                    rstEstadistica.Close
            
                End If
                
            Case Else
            
        End Select
        
        For Ciclo = 1 To Lugar
    
            For Dada = Ciclo + 1 To Lugar
    
                If Vector(Ciclo, 8) > Vector(Dada, 8) Then
    
                    Auxi1 = Vector(Ciclo, 1)
                    Auxi2 = Vector(Ciclo, 2)
                    Auxi3 = Vector(Ciclo, 3)
                    Auxi4 = Vector(Ciclo, 4)
                    Auxi5 = Vector(Ciclo, 5)
                    Auxi6 = Vector(Ciclo, 6)
                    Auxi7 = Vector(Ciclo, 7)
                    Auxi8 = Vector(Ciclo, 8)
                    Auxi9 = Vector(Ciclo, 9)
                    Auxi10 = Vector(Ciclo, 10)
                    Auxi11 = Vector(Ciclo, 11)
                    
                    Vector(Ciclo, 1) = Vector(Dada, 1)
                    Vector(Ciclo, 2) = Vector(Dada, 2)
                    Vector(Ciclo, 3) = Vector(Dada, 3)
                    Vector(Ciclo, 4) = Vector(Dada, 4)
                    Vector(Ciclo, 5) = Vector(Dada, 5)
                    Vector(Ciclo, 6) = Vector(Dada, 6)
                    Vector(Ciclo, 7) = Vector(Dada, 7)
                    Vector(Ciclo, 8) = Vector(Dada, 8)
                    Vector(Ciclo, 9) = Vector(Dada, 9)
                    Vector(Ciclo, 10) = Vector(Dada, 10)
                    Vector(Ciclo, 11) = Vector(Dada, 11)
                    
                    Vector(Dada, 1) = Auxi1
                    Vector(Dada, 2) = Auxi2
                    Vector(Dada, 3) = Auxi3
                    Vector(Dada, 4) = Auxi4
                    Vector(Dada, 5) = Auxi5
                    Vector(Dada, 6) = Auxi6
                    Vector(Dada, 7) = Auxi7
                    Vector(Dada, 8) = Auxi8
                    Vector(Dada, 9) = Auxi9
                    Vector(Dada, 10) = Auxi10
                    Vector(Dada, 11) = Auxi11
    
                End If
    
            Next Dada
    
        Next Ciclo
        
        XEmpresa = Wempresa
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        For Cicla = 1 To Lugar
        
            Renglon = Renglon + 1
                    
            WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
            WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
            WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
                            
            spCliente = "ConsultaCliente " + "'" + Vector(Cicla, 4) + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WMuestra.TextMatrix(Renglon, 4) = rstCliente!Razon
                rstCliente.Close
                    Else
                WMuestra.TextMatrix(Renglon, 4) = ""
            End If
                            
            WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
            WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
            WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
            WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
            WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
            WMuestra.TextMatrix(Renglon, 10) = ""
        
        Next Cicla
                    
        Call Conecta_Empresa
    
    End If
    
    
    
    
    
                
    Rem dada
    Rem PROCESA LAS HOJAS
    Rem dada
    
    If Tipo.ListIndex = 1 Then
        
        Erase Vector
        Lugar = 0
        
        Rem XParam = "'" + Terminado.Text + "','" _
        rem              + Terminado.Text + "'"
        Rem spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
        Sql1 = "Select Hoja.Marca, Hoja.Fecha, Hoja.Terminado, Hoja.Cantidad, Hoja.Hoja, Hoja.Lote1, Hoja.Lote2, Hoja.Lote3, Hoja.Canti1, Hoja.Canti2, Hoja.Canti3, Hoja.Tipo, Hoja.FechaFinal"
        Sql2 = " FROM Hoja"
        Sql3 = " Where Hoja.Terminado >= " + "'" + Terminado.Text + "'"
        Sql4 = " and Hoja.Terminado <= " + "'" + Terminado.Text + "'"
        Sql5 = " and Hoja.Marca <> " + "'" + "X" + "'"
        Sql6 = " and Hoja.Tipo = " + "'" + "T" + "'"
        spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            With rstHoja
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Rem  If rstHoja!Hoja = 55773 Then Stop
                    
                    XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                    If XFec < WOrdFechaCierre Then
                    
                        Else
                                
                    If rstHoja!Tipo = "T" Then
                    
                        WTerminado = rstHoja!Terminado
                        WCantidad = rstHoja!Cantidad
                        WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
                        WFechaFinal = Trim(WFechaFinal)
                        If WFechaFinal <> "" Then
                            WFecha = WFechaFinal
                                Else
                            WFecha = rstHoja!Fecha
                        End If
                        Rem WFecha = rstHoja!Fecha
                        WHoja = rstHoja!Hoja
                        
                        Erase XLote
                    
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                    
                        If XLote(1, 2) = 0 Then
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                    
                        For x = 1 To 3
                    
                            If XLote(x, 2) <> 0 Then
                        
                                If WCantidad <> 0 Then
                            
                                    WCantidad = XLote(x, 2)
                                    Lugar = Lugar + 1
                            
                                    Vector(Lugar, 1) = WFecha
                                    Vector(Lugar, 2) = "Hoja"
                                    Vector(Lugar, 3) = WHoja
                                    Vector(Lugar, 4) = ""
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WCantidad))
                                    Vector(Lugar, 7) = ""
                                    Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    Vector(Lugar, 9) = ""
                                    Vector(Lugar, 10) = XLote(x, 1)
                                    Vector(Lugar, 11) = ""
        
                                    WXSalidas = WXSalidas + WCantidad
                                    
                                End If
                                
                            End If
                            
                        Next x
                    
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
        
        For Ciclo = 1 To Lugar
    
            For Dada = Ciclo + 1 To Lugar
    
                If Vector(Ciclo, 8) > Vector(Dada, 8) Then
    
                    Auxi1 = Vector(Ciclo, 1)
                    Auxi2 = Vector(Ciclo, 2)
                    Auxi3 = Vector(Ciclo, 3)
                    Auxi4 = Vector(Ciclo, 4)
                    Auxi5 = Vector(Ciclo, 5)
                    Auxi6 = Vector(Ciclo, 6)
                    Auxi7 = Vector(Ciclo, 7)
                    Auxi8 = Vector(Ciclo, 8)
                    Auxi9 = Vector(Ciclo, 9)
                    Auxi10 = Vector(Ciclo, 10)
                    Auxi11 = Vector(Ciclo, 11)
                    
                    Vector(Ciclo, 1) = Vector(Dada, 1)
                    Vector(Ciclo, 2) = Vector(Dada, 2)
                    Vector(Ciclo, 3) = Vector(Dada, 3)
                    Vector(Ciclo, 4) = Vector(Dada, 4)
                    Vector(Ciclo, 5) = Vector(Dada, 5)
                    Vector(Ciclo, 6) = Vector(Dada, 6)
                    Vector(Ciclo, 7) = Vector(Dada, 7)
                    Vector(Ciclo, 8) = Vector(Dada, 8)
                    Vector(Ciclo, 9) = Vector(Dada, 9)
                    Vector(Ciclo, 10) = Vector(Dada, 10)
                    Vector(Ciclo, 11) = Vector(Dada, 11)
                    
                    Vector(Dada, 1) = Auxi1
                    Vector(Dada, 2) = Auxi2
                    Vector(Dada, 3) = Auxi3
                    Vector(Dada, 4) = Auxi4
                    Vector(Dada, 5) = Auxi5
                    Vector(Dada, 6) = Auxi6
                    Vector(Dada, 7) = Auxi7
                    Vector(Dada, 8) = Auxi8
                    Vector(Dada, 9) = Auxi9
                    Vector(Dada, 10) = Auxi10
                    Vector(Dada, 11) = Auxi11
    
                End If
    
            Next Dada
    
        Next Ciclo
        
        For Cicla = 1 To Lugar
        
            Renglon = Renglon + 1
                    
            WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
            WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
            WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
            WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
            WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
            WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
            WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
            WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
            WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
            WMuestra.TextMatrix(Renglon, 10) = ""
        
        Next Cicla
    
    End If
    
    
    
    
    Rem dada
    Rem PROCESA LAS HOJAS
    Rem dada
    
    Erase Vector
    Lugar = 0
    
        XParam = "'" + Terminado.Text + "','" _
                     + Terminado.Text + "'"
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
                            
                                WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                Call Redondeo(WSaldo)
                                If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
                                 
                                    WProducto = rstHoja!Producto
                                    WCantidad = rstHoja!Real
                                    WCantidadII = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
                                    WCantidadIII = WCantidad + WCantidadII
                                    
                                    WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
                                    WFechaFinal = Trim(WFechaFinal)
                                    If WFechaFinal <> "" Then
                                        WFecha = WFechaFinal
                                            Else
                                        WFecha = rstHoja!Fecha
                                    End If
                                    Rem WFecha = rstHoja!Fecha
                                    
                                    WHoja = rstHoja!Hoja
                                    WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                    Call Redondeo(WSaldo)
                                    WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            
                                    If WCantidad <> 0 Or WCantidadII <> 0 Then
                                        
                                        aa = rstHoja!Marca
                                        aa = rstHoja!Cantidad
                                        Lugar = Lugar + 1
                                        
                                        Vector(Lugar, 1) = WFecha
                                        Vector(Lugar, 2) = "Hoja"
                                        Vector(Lugar, 3) = WHoja
                                        Vector(Lugar, 4) = ""
                                        Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(WCantidad))
                                        Vector(Lugar, 6) = ""
                                        Vector(Lugar, 7) = ""
                                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        Vector(Lugar, 10) = WHoja
                                        Vector(Lugar, 11) = Pusing("###,###,###.##", Str$(WSaldo))
                                        Vector(Lugar, 12) = WMarcaVencida
                                        
                                        WXEntradas = WXEntradas + WCantidad
                                        
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(Dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                Auxi12 = Vector(Ciclo, 12)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                Vector(Ciclo, 11) = Vector(Dada, 11)
                Vector(Ciclo, 12) = Vector(Dada, 12)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10
                Vector(Dada, 11) = Auxi11
                Vector(Dada, 12) = Auxi12

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
        WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
        WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
        WMuestra.TextMatrix(Renglon, 10) = Vector(Cicla, 12)
    
    Next Cicla
    
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    Rem dada
    
    
    
    
    If Tipo.ListIndex = 1 Then
        
        Erase Vector
        Lugar = 0
        
        Rem XParam = "'" + Terminado.Text + "','" _
        Rem              + Terminado.Text + "'"
        Rem spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
        Sql1 = "Select Movvar.Marca, Movvar.Tipo, Movvar.Terminado, Movvar.Cantidad, Movvar.Fecha, Movvar.Codigo, Movvar.Movi, Movvar.Lote, Movvar.TipoMov, Movvar.Observaciones"
        Sql2 = " FROM Movvar"
        Sql3 = " Where Movvar.Terminado >= " + "'" + Terminado.Text + "'"
        Sql4 = " and Movvar.Terminado <= " + "'" + Terminado.Text + "'"
        Sql5 = " and Movvar.Marca <> " + "'" + "X" + "'"
        Sql6 = " and Movvar.Tipo = " + "'" + "T" + "'"
        spMovvar = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
        
            With rstMovvar
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstMovvar!Tipo = "T" Then
                    
                        WTerminado = rstMovvar!Terminado
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = WFecha
                        If Val(rstMovvar!Tipomov) = 0 Or Val(rstMovvar!Tipomov) = 1 Then
                            Vector(Lugar, 2) = "Mov.Var"
                                Else
                            Vector(Lugar, 2) = "Guia In"
                        End If
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovvar!Observaciones
                        If WMovi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(WCantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + WCantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WCantidad))
                            WXSalidas = WXSalidas + WCantidad
                        End If
                        Vector(Lugar, 7) = ""
                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        Vector(Lugar, 10) = WLote
                        Vector(Lugar, 11) = ""
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                
                End If
                    
            End With
            rstMovvar.Close
        End If
        
        For Ciclo = 1 To Lugar
    
            For Dada = Ciclo + 1 To Lugar
    
                If Vector(Ciclo, 8) > Vector(Dada, 8) Then
    
                    Auxi1 = Vector(Ciclo, 1)
                    Auxi2 = Vector(Ciclo, 2)
                    Auxi3 = Vector(Ciclo, 3)
                    Auxi4 = Vector(Ciclo, 4)
                    Auxi5 = Vector(Ciclo, 5)
                    Auxi6 = Vector(Ciclo, 6)
                    Auxi7 = Vector(Ciclo, 7)
                    Auxi8 = Vector(Ciclo, 8)
                    Auxi9 = Vector(Ciclo, 9)
                    Auxi10 = Vector(Ciclo, 10)
                    Auxi11 = Vector(Ciclo, 11)
                    
                    Vector(Ciclo, 1) = Vector(Dada, 1)
                    Vector(Ciclo, 2) = Vector(Dada, 2)
                    Vector(Ciclo, 3) = Vector(Dada, 3)
                    Vector(Ciclo, 4) = Vector(Dada, 4)
                    Vector(Ciclo, 5) = Vector(Dada, 5)
                    Vector(Ciclo, 6) = Vector(Dada, 6)
                    Vector(Ciclo, 7) = Vector(Dada, 7)
                    Vector(Ciclo, 8) = Vector(Dada, 8)
                    Vector(Ciclo, 9) = Vector(Dada, 9)
                    Vector(Ciclo, 10) = Vector(Dada, 10)
                    Vector(Ciclo, 11) = Vector(Dada, 11)
                    
                    Vector(Dada, 1) = Auxi1
                    Vector(Dada, 2) = Auxi2
                    Vector(Dada, 3) = Auxi3
                    Vector(Dada, 4) = Auxi4
                    Vector(Dada, 5) = Auxi5
                    Vector(Dada, 6) = Auxi6
                    Vector(Dada, 7) = Auxi7
                    Vector(Dada, 8) = Auxi8
                    Vector(Dada, 9) = Auxi9
                    Vector(Dada, 10) = Auxi10
                    Vector(Dada, 11) = Auxi11
    
                End If
    
            Next Dada
    
        Next Ciclo
        
        For Cicla = 1 To Lugar
        
            Renglon = Renglon + 1
                    
            WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
            WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
            WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
            WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
            WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
            WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
            WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
            WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
            WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
            WMuestra.TextMatrix(Renglon, 10) = ""
        
        Next Cicla
        
    End If
    
    
    
    Rem dada
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNO
    Rem dada
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
                   
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
                        
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            Call Redondeo(WSaldo)
                            If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
                            
                                WTerminado = rstMovguia!Terminado
                                WCantidad = rstMovguia!Cantidad
                                WFecha = rstMovguia!Fecha
                                WCodigo = rstMovguia!Codigo
                                WMovi = rstMovguia!Movi
                                WDestino = rstMovguia!Destino
                                WTipomov = rstMovguia!Tipomov
                                WLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                WPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
            
                                Lugar = Lugar + 1
                                
                                Vector(Lugar, 1) = WFecha
                                If Val(WCodigo) > 900000 Then
                                    Vector(Lugar, 2) = "Prestamo"
                                    Vector(Lugar, 3) = WCodigo - 900000
                                        Else
                                    Vector(Lugar, 2) = "Guia In"
                                    Vector(Lugar, 3) = WCodigo
                                End If
                                Rem Vector(Lugar, 4) = rstMovguia!Observaciones
                                
                                If WMovi = "E" Then
                                
                                    Select Case WTipomov
                                        Case 1
                                            Vector(Lugar, 4) = "Recepcion de Surfactan"
                                        Case 2
                                            Vector(Lugar, 4) = "Recepcion de Pellital"
                                        Case 3
                                            Vector(Lugar, 4) = "Recepcion de Surfactan II"
                                        Case 4
                                            Vector(Lugar, 4) = "Recepcion de Pellital II"
                                        Case 5
                                            Vector(Lugar, 4) = "Recepcion de Surfactan III"
                                        Case 6
                                            Vector(Lugar, 4) = "Recepcion de Surfactan IV"
                                        Case 7
                                            Vector(Lugar, 4) = "Recepcion de Surfactan V"
                                        Case 8
                                            Vector(Lugar, 4) = "Recepcion de Pellital V"
                                        Case 9
                                            If Left$(Terminado.Text, 2) = "DW" Then
                                                Vector(Lugar, 2) = "Ajuste"
                                                Vector(Lugar, 4) = "Trasformacion a MP"
                                                    Else
                                                Vector(Lugar, 4) = "Recepcion de Pellital IV"
                                            End If
                                        Case 10
                                            Vector(Lugar, 4) = "Recepcion de Surfactan VI"
                                        Case 11
                                            Vector(Lugar, 4) = "Recepcion de Surfactan VII"
                                        Case Else
                                    End Select
                                    Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(WCantidad))
                                    Vector(Lugar, 6) = ""
                                    Vector(Lugar, 10) = WLote
                                    Vector(Lugar, 11) = Pusing("###,###,###.##", Str$(WSaldo))
                                    Vector(Lugar, 12) = WMarcaVencida
                                    WXEntradas = WXEntradas + WCantidad
                                    
                                        Else
                                        
                                    Select Case WDestino
                                        Case 1
                                            Vector(Lugar, 4) = "Envio a Surfactan"
                                        Case 2
                                            Vector(Lugar, 4) = "Envio a Pellital"
                                        Case 3
                                            Vector(Lugar, 4) = "Envio a Surfactan II"
                                        Case 4
                                            Vector(Lugar, 4) = "Envio a Pellital II"
                                        Case 5
                                            Vector(Lugar, 4) = "Envio a Surfactan III"
                                        Case 6
                                            Vector(Lugar, 4) = "Envio a Surfactan IV"
                                        Case 7
                                            Vector(Lugar, 4) = "Envio a Surfactan V"
                                        Case 8
                                            Vector(Lugar, 4) = "Envio a Pellital V"
                                        Case 9
                                            If Left$(Terminado.Text, 2) = "DW" Then
                                                Vector(Lugar, 2) = "Ajuste"
                                                Vector(Lugar, 4) = "Trasformacion a MP"
                                                    Else
                                                Vector(Lugar, 4) = "Envio a Pellital IV"
                                            End If
                                        Case 10
                                            Vector(Lugar, 4) = "Envio a Surfactan VI"
                                        Case 11
                                            Vector(Lugar, 4) = "Envio a Surfactan VII"
                                        Case Else
                                    End Select
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WCantidad))
                                    Vector(Lugar, 10) = WPartida
                                    Vector(Lugar, 11) = ""
                                    Vector(Lugar, 12) = WMarcaVencida
                                    WXSalidas = WXSalidas + WCantidad
                                    
                                End If
                                Vector(Lugar, 7) = ""
                                Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(Dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                Auxi12 = Vector(Ciclo, 12)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                Vector(Ciclo, 11) = Vector(Dada, 11)
                Vector(Ciclo, 12) = Vector(Dada, 12)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10
                Vector(Dada, 11) = Auxi11
                Vector(Dada, 12) = Auxi12

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
        WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
        WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
        WMuestra.TextMatrix(Renglon, 10) = Vector(Cicla, 12)
    
    Next Cicla
    
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    Rem dada
    
    If Tipo.ListIndex = 1 Then
    
        Erase Vector
        Lugar = 0
        
        XParam = "'" + Terminado.Text + "','" _
                     + Terminado.Text + "'"
        spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovlab.RecordCount > 0 Then
        
            With rstMovlab
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstMovlab!Marca = "X" Then
                    
                        Else
                    
                    If rstMovlab!Tipo = "T" Then
                    
                        WTerminado = rstMovlab!Terminado
                        WCantidad = rstMovlab!Cantidad
                        WFecha = rstMovlab!Fecha
                        WCodigo = rstMovlab!Codigo
                        WMovi = rstMovlab!Movi
                        WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
    
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = WFecha
                        Vector(Lugar, 2) = "Mov.Lab"
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovlab!Observaciones
                        
                        If WMovi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(WCantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + WCantidad
                                    Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WCantidad))
                            WXSalidas = WXSalidas + WCantidad
                        End If
                        Vector(Lugar, 7) = ""
                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        Vector(Lugar, 10) = WLote
                        Vector(Lugar, 11) = ""
                        
                    
                    End If
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                
                End If
                
            End With
            rstMovlab.Close
        End If
        
        For Ciclo = 1 To Lugar
    
            For Dada = Ciclo + 1 To Lugar
    
                If Vector(Ciclo, 8) > Vector(Dada, 8) Then
    
                    Auxi1 = Vector(Ciclo, 1)
                    Auxi2 = Vector(Ciclo, 2)
                    Auxi3 = Vector(Ciclo, 3)
                    Auxi4 = Vector(Ciclo, 4)
                    Auxi5 = Vector(Ciclo, 5)
                    Auxi6 = Vector(Ciclo, 6)
                    Auxi7 = Vector(Ciclo, 7)
                    Auxi8 = Vector(Ciclo, 8)
                    Auxi9 = Vector(Ciclo, 9)
                    Auxi10 = Vector(Ciclo, 10)
                    Auxi11 = Vector(Ciclo, 11)
                    
                    Vector(Ciclo, 1) = Vector(Dada, 1)
                    Vector(Ciclo, 2) = Vector(Dada, 2)
                    Vector(Ciclo, 3) = Vector(Dada, 3)
                    Vector(Ciclo, 4) = Vector(Dada, 4)
                    Vector(Ciclo, 5) = Vector(Dada, 5)
                    Vector(Ciclo, 6) = Vector(Dada, 6)
                    Vector(Ciclo, 7) = Vector(Dada, 7)
                    Vector(Ciclo, 8) = Vector(Dada, 8)
                    Vector(Ciclo, 9) = Vector(Dada, 9)
                    Vector(Ciclo, 10) = Vector(Dada, 10)
                    Vector(Ciclo, 11) = Vector(Dada, 11)
                    
                    Vector(Dada, 1) = Auxi1
                    Vector(Dada, 2) = Auxi2
                    Vector(Dada, 3) = Auxi3
                    Vector(Dada, 4) = Auxi4
                    Vector(Dada, 5) = Auxi5
                    Vector(Dada, 6) = Auxi6
                    Vector(Dada, 7) = Auxi7
                    Vector(Dada, 8) = Auxi8
                    Vector(Dada, 9) = Auxi9
                    Vector(Dada, 10) = Auxi10
                    Vector(Dada, 11) = Auxi11
    
                End If
    
            Next Dada
    
        Next Ciclo
        
        For Cicla = 1 To Lugar
        
            Renglon = Renglon + 1
                    
            WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
            WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
            WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
            WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
            WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
            WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
            WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
            WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
            WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
            WMuestra.TextMatrix(Renglon, 10) = ""
        
        Next Cicla
        
    End If
        
    
    
    
    
    Rem dada
    Rem REMITOS EN CONSIGNACION
    Rem dada
    
    If Tipo.ListIndex = 1 Then
    
        Erase Vector
        Lugar = 0
        
        XParam = "'" + Terminado.Text + "','" _
                     + Terminado.Text + "'"
        spConsig = "ListaConsigTerminado" + XParam
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount > 0 Then
        
            With rstConsig
        
                .MoveFirst
                
                If .NoMatch = False Then
                
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstConsig!Marca <> "X" Then
                    
                    WTerminado = rstConsig!Terminado
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WFecha = rstConsig!Fecha
                    WCodigo = rstConsig!Numero
                    WLote = IIf(IsNull(rstConsig!Lote), "", rstConsig!Lote)
                        
                    If WCantidad <> 0 Then
                    
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = WFecha
                        Vector(Lugar, 2) = "Rem.Con."
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstConsig!Observaciones
                        Vector(Lugar, 5) = ""
                        Vector(Lugar, 6) = Pusing("###,###,###.##", Str$(WCantidad))
                        WXSalidas = WXSalidas + WCantidad
                        Vector(Lugar, 7) = ""
                        Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        Vector(Lugar, 9) = rstConsig!Cliente
                        Vector(Lugar, 10) = WLote
                        Vector(Lugar, 11) = ""
                    
                    End If
                    
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                
                End If
                    
            End With
            rstConsig.Close
        End If
        
        For Ciclo = 1 To Lugar
        
            WImpre1 = Vector(Ciclo, 9)
            
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WImpre2 = rstCliente!Razon
                rstCliente.Close
                    Else
                WImpre2 = ""
            End If
            
            Call Conecta_Empresa
        
            Vector(Ciclo, 4) = WImpre2
        
        Next Ciclo
        
        For Ciclo = 1 To Lugar
    
            For Dada = Ciclo + 1 To Lugar
    
                If Vector(Ciclo, 8) > Vector(Dada, 8) Then
    
                    Auxi1 = Vector(Ciclo, 1)
                    Auxi2 = Vector(Ciclo, 2)
                    Auxi3 = Vector(Ciclo, 3)
                    Auxi4 = Vector(Ciclo, 4)
                    Auxi5 = Vector(Ciclo, 5)
                    Auxi6 = Vector(Ciclo, 6)
                    Auxi7 = Vector(Ciclo, 7)
                    Auxi8 = Vector(Ciclo, 8)
                    Auxi9 = Vector(Ciclo, 9)
                    Auxi10 = Vector(Ciclo, 10)
                    Auxi11 = Vector(Ciclo, 11)
                    
                    Vector(Ciclo, 1) = Vector(Dada, 1)
                    Vector(Ciclo, 2) = Vector(Dada, 2)
                    Vector(Ciclo, 3) = Vector(Dada, 3)
                    Vector(Ciclo, 4) = Vector(Dada, 4)
                    Vector(Ciclo, 5) = Vector(Dada, 5)
                    Vector(Ciclo, 6) = Vector(Dada, 6)
                    Vector(Ciclo, 7) = Vector(Dada, 7)
                    Vector(Ciclo, 8) = Vector(Dada, 8)
                    Vector(Ciclo, 9) = Vector(Dada, 9)
                    Vector(Ciclo, 10) = Vector(Dada, 10)
                    Vector(Ciclo, 11) = Vector(Dada, 11)
                    
                    Vector(Dada, 1) = Auxi1
                    Vector(Dada, 2) = Auxi2
                    Vector(Dada, 3) = Auxi3
                    Vector(Dada, 4) = Auxi4
                    Vector(Dada, 5) = Auxi5
                    Vector(Dada, 6) = Auxi6
                    Vector(Dada, 7) = Auxi7
                    Vector(Dada, 8) = Auxi8
                    Vector(Dada, 9) = Auxi9
                    Vector(Dada, 10) = Auxi10
                    Vector(Dada, 11) = Auxi11
    
                End If
    
            Next Dada
    
        Next Ciclo
        
        For Cicla = 1 To Lugar
        
            Renglon = Renglon + 1
                    
            WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
            WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
            WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
            WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
            WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
            WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
            WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
            WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
            WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
            WMuestra.TextMatrix(Renglon, 10) = ""
        
        Next Cicla
        
    End If
        
        
    Rem dada
    Rem PROCESA LOS las devoluciones de mercaderia
    Rem dada
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Terminado.Text + "','" _
                 + Terminado.Text + "'"
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
                    
                    WCodigo = rstEntdev!Codigo
                    
                    If rstEntdev!Marca = "X" Then
                    
                                Else
                        
                        WTerminado = rstEntdev!Terminado
                        WCantidad = rstEntdev!Cantidad
                        WLaboratorio = rstEntdev!Laboratorio
                        WFecha = rstEntdev!Fecha
                        WCodigo = rstEntdev!Codigo
                        WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                        WSaldo = rstEntdev!Saldo
                            
                        If Tipo.ListIndex = 1 Or WSaldo <> 0 Then
                                
                            Lugar = Lugar + 1
                                
                            Vector(Lugar, 1) = WFecha
                            Vector(Lugar, 2) = "Ent.Dev"
                            Vector(Lugar, 3) = WCodigo
                            Vector(Lugar, 4) = rstEntdev!Observaciones
                            Vector(Lugar, 5) = Pusing("###,###,###.##", Str$(WCantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + WCantidad
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            Vector(Lugar, 10) = WLote
                            Vector(Lugar, 11) = Pusing("###,###,###.##", Str$(WSaldo))
                            
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
    
    For Ciclo = 1 To Lugar

        For Dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 8) > Vector(Dada, 8) Then

                Auxi1 = Vector(Ciclo, 1)
                Auxi2 = Vector(Ciclo, 2)
                Auxi3 = Vector(Ciclo, 3)
                Auxi4 = Vector(Ciclo, 4)
                Auxi5 = Vector(Ciclo, 5)
                Auxi6 = Vector(Ciclo, 6)
                Auxi7 = Vector(Ciclo, 7)
                Auxi8 = Vector(Ciclo, 8)
                Auxi9 = Vector(Ciclo, 9)
                Auxi10 = Vector(Ciclo, 10)
                Auxi11 = Vector(Ciclo, 11)
                
                Vector(Ciclo, 1) = Vector(Dada, 1)
                Vector(Ciclo, 2) = Vector(Dada, 2)
                Vector(Ciclo, 3) = Vector(Dada, 3)
                Vector(Ciclo, 4) = Vector(Dada, 4)
                Vector(Ciclo, 5) = Vector(Dada, 5)
                Vector(Ciclo, 6) = Vector(Dada, 6)
                Vector(Ciclo, 7) = Vector(Dada, 7)
                Vector(Ciclo, 8) = Vector(Dada, 8)
                Vector(Ciclo, 9) = Vector(Dada, 9)
                Vector(Ciclo, 10) = Vector(Dada, 10)
                Vector(Ciclo, 11) = Vector(Dada, 11)
                
                Vector(Dada, 1) = Auxi1
                Vector(Dada, 2) = Auxi2
                Vector(Dada, 3) = Auxi3
                Vector(Dada, 4) = Auxi4
                Vector(Dada, 5) = Auxi5
                Vector(Dada, 6) = Auxi6
                Vector(Dada, 7) = Auxi7
                Vector(Dada, 8) = Auxi8
                Vector(Dada, 9) = Auxi9
                Vector(Dada, 10) = Auxi10
                Vector(Dada, 11) = Auxi11

            End If

        Next Dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WMuestra.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WMuestra.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WMuestra.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WMuestra.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WMuestra.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WMuestra.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WMuestra.TextMatrix(Renglon, 7) = Vector(Cicla, 7)
        WMuestra.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
        WMuestra.TextMatrix(Renglon, 9) = Vector(Cicla, 11)
        WMuestra.TextMatrix(Renglon, 10) = ""
    
    Next Cicla
    
    ZZFechaActual = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
    
    For Cicla = 1 To Renglon
    
        ZZFecha = WMuestra.TextMatrix(Cicla, 1)
        ZZTipo = WMuestra.TextMatrix(Cicla, 2)
        ZZNumero = WMuestra.TextMatrix(Cicla, 3)
        ZZObservaciones = WMuestra.TextMatrix(Cicla, 4)
        ZZEntrada = WMuestra.TextMatrix(Cicla, 5)
        ZZSalida = WMuestra.TextMatrix(Cicla, 6)
        ZZRemito = WMuestra.TextMatrix(Cicla, 7)
        ZZPartida = WMuestra.TextMatrix(Cicla, 8)
        ZZSaldo = WMuestra.TextMatrix(Cicla, 9)
        ZZMarcaVencida = WMuestra.TextMatrix(Cicla, 10)
        
        If Val(ZZSaldo) <> 0 And ZZMarcaVencida = "S" Then
            
             WMuestra.Row = Cicla
             
             WMuestra.Col = 1
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 2
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 3
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 4
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 5
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 6
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 7
             WMuestra.CellBackColor = &HC0FFFF
            
             WMuestra.Col = 8
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 9
             WMuestra.CellBackColor = &HC0FFFF
             
             WMuestra.Col = 10
             WMuestra.CellBackColor = &HC0FFFF
            
        End If
        
        Rem BY NAN
        If Val(ZZSaldo) <> 0 And ZZMarcaVencida = "V" Then
            
             WMuestra.Row = Cicla
             
             WMuestra.Col = 1
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 2
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 3
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 4
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 5
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 6
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 7
             WMuestra.CellBackColor = &HFF&
            
             WMuestra.Col = 8
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 9
             WMuestra.CellBackColor = &HFF&
             
             WMuestra.Col = 10
             WMuestra.CellBackColor = &HFF&
            
        End If
    
    Next Cicla
    
    WXStock = WXInicial + WXEntradas - WXSalidas
    
    XInicial.Text = Pusing("###,###,###.##", Str$(WXInicial))
    XEntradas.Text = Pusing("###,###,###.##", Str$(WXEntradas))
    XSalidas.Text = Pusing("###,###,###.##", Str$(WXSalidas))
    XStock.Text = Pusing("###,###,###.##", Str$(WXStock))
    
    spTerminado = "ConsultaTerminado " + "'" + "NK" + Right$(Terminado.Text, 10) + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Nk.Text = Pusing("###,###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        Nk.Text = "0.00"
    End If
    
    spTerminado = "ConsultaTerminado " + "'" + "RE" + Right$(Terminado.Text, 10) + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        RE.Text = Pusing("###,###,###.##", rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas)
        rstTerminado.Close
            Else
        RE.Text = "0.00"
    End If
    
    WMuestra.TopRow = 1
    WMuestra.Row = 1
    WMuestra.Col = 1
    
    PrgConsFicTer.WindowState = 0

    Exit Sub
    
Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    StockCons.Visible = False
    Resume Next

End Sub

Private Sub listpedido_Click()

    Label11.Caption = "0"
    Label11.Caption = Pusing("###,###,###.##", Label11.Caption)
    Label15.Caption = "0"
    Label15.Caption = Pusing("###,###,###.##", Label15.Caption)
    
    Total = 0
   
    Empres = Wempresa
    
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Case 2, 4, 8, 9
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Case Else
    End Select
   
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
 
    Sql1 = "Select Pedido.Cantidad,Pedido.Facturado"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " and Pedido.facturado < Pedido.Cantidad "
    sptotalpt = Sql1 + Sql2 + Sql3 + Sql4
    Set rsttotalpt = db.OpenRecordset(sptotalpt, dbOpenSnapshot, dbSQLPassThrough)
    If rsttotalpt.RecordCount > 0 Then
    
        With rsttotalpt
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
            
                    Total = Total + rsttotalpt!Cantidad - rsttotalpt!Facturado
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
        
            End If
        End With
        rsttotalpt.Close
        
        Label11.Caption = Str$(Total)
        Label11.Caption = Pusing("###,###,###.##", Label11.Caption)
        
    End If

    
    XEmpresa = Wempresa
    Call Conecta_Empresa

    ZStock = Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(WStock6.Caption) + Val(WStock7.Caption) + Val(StkProceso.Text)
    ZFaltante = Val(Label11.Caption) - ZStock

    If ZFaltante <= 0 Then
        Label15.Caption = "0"
            Else
        Label15.Caption = Str$(ZFaltante)
    End If
    Label15.Caption = Pusing("###,###,###.##", Label15.Caption)
    
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        Terminado.Text = WTerminado
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            Unidad.Text = rstTerminado!Unidad
            Deposito.Text = rstTerminado!Deposito
            StkProceso.Text = Pusing("###,###,###.##", Str$(rstTerminado!Proceso))
            rstTerminado.Close
            Call Proceso_Click
            Call listpedido_Click
                Else
            Terminado.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Vector()

    WMuestra.Clear
    WMuestra.Font.Bold = True
    
    WMuestra.FixedCols = 1
    WMuestra.Cols = 11
    WMuestra.FixedRows = 1
    WMuestra.Rows = 2001
    
    WMuestra.ColWidth(0) = 200
    WMuestra.Row = 0
    For Ciclo = 1 To WMuestra.Cols - 1
        WMuestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                WMuestra.Text = "Fecha"
                WMuestra.ColWidth(Ciclo) = 1200
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WMuestra.Text = "Tipo"
                WMuestra.ColWidth(Ciclo) = 900
                WMuestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WMuestra.Text = "Numero"
                WMuestra.ColWidth(Ciclo) = 1000
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WMuestra.Text = "Observaciones"
                WMuestra.ColWidth(Ciclo) = 2600
                WMuestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WMuestra.Text = "Entradas"
                WMuestra.ColWidth(Ciclo) = 1000
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WMuestra.Text = "Salidas"
                WMuestra.ColWidth(Ciclo) = 1000
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WMuestra.Text = "Remito"
                WMuestra.ColWidth(Ciclo) = 900
                WMuestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WMuestra.Text = "Partida"
                WMuestra.ColWidth(Ciclo) = 900
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WMuestra.Text = "Saldo"
                WMuestra.ColWidth(Ciclo) = 1000
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WMuestra.Text = ""
                WMuestra.ColWidth(Ciclo) = 10
                WMuestra.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WMuestra.Row = 0
    For Ciclo = 1 To WMuestra.Cols - 1
        WMuestra.Col = Ciclo
        WTitulo(Ciclo).Text = WMuestra.Text
        WTitulo(Ciclo).Left = WMuestra.CellLeft + WMuestra.Left
        WTitulo(Ciclo).Top = WMuestra.CellTop + WMuestra.Top
        WTitulo(Ciclo).Width = WMuestra.CellWidth
        WTitulo(Ciclo).Height = WMuestra.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WMuestra.Cols - 1
        WAncho = WAncho + WMuestra.ColWidth(Ciclo)
    Next Ciclo
    WMuestra.Width = WAncho

    ' Size the columns.
    Font.Name = WMuestra.Font.Name
    Font.Size = WMuestra.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WMuestra.AllowUserResizing = flexResizeBoth
    
    WMuestra.Col = 1
    WMuestra.Row = 1
    
End Sub









