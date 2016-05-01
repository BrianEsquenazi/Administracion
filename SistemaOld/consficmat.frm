VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsFicMat 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Ficha de Stock Materia Prima"
   ClientHeight    =   7665
   ClientLeft      =   405
   ClientTop       =   720
   ClientWidth     =   11385
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   11385
   Begin VB.Frame PantaHistoria 
      Caption         =   "Historial"
      Height          =   3135
      Left            =   120
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox HistoriaCarpeta 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton PantaHistoriaCierra 
         Caption         =   "Cierra"
         Height          =   495
         Left            =   3360
         TabIndex        =   56
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox HistoriaFactura 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   " "
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox HistoriaRemito 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox HistoriaInforme 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox HistoriaOrden 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin MSMask.MaskEdBox HistoriaFechaOrden 
         Height          =   285
         Left            =   5280
         TabIndex        =   47
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox HistoriaFechaInforme 
         Height          =   285
         Left            =   5280
         TabIndex        =   50
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox HistoriaFechaFactura 
         Height          =   285
         Left            =   5280
         TabIndex        =   55
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label11 
         Caption         =   "Carpeta"
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
         Left            =   840
         TabIndex        =   58
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Factura"
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
         Left            =   840
         TabIndex        =   53
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label9 
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
         Left            =   840
         TabIndex        =   51
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Informe de Recepcion"
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
         Left            =   840
         TabIndex        =   48
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Orden de Compra"
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
         Left            =   840
         TabIndex        =   45
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Historial 
      Caption         =   "Historial"
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
      TabIndex        =   43
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Solicitud 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1280
      Width           =   1335
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
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3000
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3000
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3000
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3000
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3000
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3000
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3000
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3000
      Width           =   375
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11160
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WLotemat.rpt"
   End
   Begin VB.Frame StockCons 
      Caption         =   "Stock Consolidado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7800
      TabIndex        =   16
      Top             =   0
      Width           =   2415
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
         Left            =   1200
         TabIndex        =   42
         Top             =   1680
         Width           =   975
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
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   975
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
         Left            =   1200
         TabIndex        =   40
         Top             =   1440
         Width           =   975
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
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   975
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
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   975
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
         Left            =   1200
         TabIndex        =   25
         Top             =   1200
         Width           =   975
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
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   960
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
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   975
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
         TabIndex        =   22
         Top             =   720
         Width           =   975
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
         TabIndex        =   21
         Top             =   480
         Width           =   975
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
         TabIndex        =   20
         Top             =   240
         Width           =   975
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
         Height          =   375
         Left            =   120
         TabIndex        =   19
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
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
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   " "
      Top             =   960
      Width           =   1335
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   " "
      Top             =   660
      Width           =   1335
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " "
      Top             =   360
      Width           =   1335
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " "
      Top             =   80
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
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
      ItemData        =   "consficmat.frx":0000
      Left            =   120
      List            =   "consficmat.frx":0007
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5295
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9340
      _Version        =   327680
      BackColor       =   16777152
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
      Height          =   1935
      Left            =   8160
      Picture         =   "consficmat.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Solicitud"
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
      TabIndex        =   38
      Top             =   1280
      Width           =   2055
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
      Top             =   960
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
      Top             =   660
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
      Top             =   360
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
      Top             =   80
      Width           =   1695
   End
   Begin VB.Label DesArticulo 
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Articulo"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsFicMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private Vector(5000, 10) As String
Private XLote(100, 7) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim XParam As String
Dim WReventa As Integer

Private WOrden As String
Private WXInicial As Double
Private WXSalidas As Double
Private WXEntradas As Double
Private WXStock As Double
Private WCanti As Double
Private WSaldo As Double
Private NombreEmpresa As String
Private WGrilla(100, 10) As String

Private Sub cmdClose_Click()
    Articulo.SetFocus
    PrgConsFicMat.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    WDesdeArticulo = "AA-000-000"
    WHastaArticulo = "ZZ-999-999"
    
    XParam = "'" + WDesdeArticulo + "','" _
                 + WHastaArticulo + "'"
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
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
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                        
                    If rstHoja!Tipo = "M" Then
                    
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        WSalidas = 0
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) = "" Then
                                XLote(Da, 2) = "0"
                            End If
                        
                            WCanti = XLote(Da, 2)
                            If WCanti <> 0 Then
                                WSalidas = WSalidas + WCanti
                            End If
                        Next Da
                        
                        If WSalidas <> rstHoja!Cantidad Then
                            aa = rstHoja!Hoja
                            AA1 = rstHoja!Articulo
                        End If
                        
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstHoja!Articulo > WArticulo Then
                Rem     Exit Do
                Rem End If
                
            Loop
            End If
        
        End With
        
        rstHoja.Close
        
    End If



End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    
    Pantalla.Clear
    WIndice.Clear

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
            
    Pantalla.Visible = True

End Sub




Private Sub Historial_Click()

    WVector1.Col = 2
    WMovimiento = WVector1.Text
    
    If WMovimiento = "Laudo" Then
        
        WVector1.Col = 3
        WLaudo = WVector1.Text
        
        HistoriaOrden.Text = ""
        HistoriaInforme.Text = ""
        HistoriaRemito.Text = ""
        HistoriaFactura.Text = ""
        HistoriaCarpeta.Text = ""
        
        HistoriaFechaOrden.Text = "  /  /    "
        HistoriaFechaInforme.Text = "  /  /    "
        HistoriaFechaFactura.Text = "  /  /    "
        
        WOrden = ""
        WInforme = ""
        WRemito = ""
        WFactura = ""
        WProveedor = ""
        WCarpeta = ""
        
        WFechaOrden = "  /  /    "
        WFechaInforme = "  /  /    "
        WFechaFactura = "  /  /    "
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + WLaudo + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + Articulo.Text + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WInforme = Str$(rstLaudo!Informe)
            WOrden = rstLaudo!Orden
            rstLaudo.Close
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Informe.Informe = " + "'" + WInforme + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                WRemito = Str$(rstInforme!Remito)
                WFechaInforme = rstInforme!Fecha
                rstInforme.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Orden.Orden = " + "'" + WOrden + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WFechaOrden = rstOrden!Fecha
                WProveedor = rstOrden!Proveedor
                WCarpeta = rstOrden!Carpeta
                rstOrden.Close
            End If
            
            XEmpresa = WEmpresa
            
            
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM IvaComp"
            ZSql = ZSql + " Where IvaComp.Remito LIKE " + "'" + "%" + Trim(WRemito) + "%" + "'"
            ZSql = ZSql + " and IvaComp.Proveedor = " + "'" + WProveedor + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                WFactura = Str$(Val(rstIvaComp!Numero))
                WFechaFactura = rstIvaComp!Fecha
                rstIvaComp.Close
            End If
            
            Call Conecta_Empresa
        
        End If
        
        HistoriaOrden.Text = WOrden
        HistoriaInforme.Text = WInforme
        HistoriaRemito.Text = WRemito
        HistoriaFactura.Text = WFactura
        HistoriaCarpeta.Text = WCarpeta
        
        HistoriaFechaOrden.Text = WFechaOrden
        HistoriaFechaInforme.Text = WFechaInforme
        HistoriaFechaFactura.Text = WFechaFactura
        
        PantaHistoria.Visible = True
        
    End If
    
End Sub

Private Sub PantaHistoriaCierra_Click()
    PantaHistoria.Visible = False
End Sub

Private Sub WVector1_DblClick()

    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
        rstArticulo.Close
    End If

    ZTipo = Left$(Articulo.Text, 2)
    If WReventa = 1 Then
        ZTipo = "DY"
    End If

    Select Case ZTipo
        Case "DY", "DS", "DQ"
        
            WVector1.Col = 7
            WPartiOri = WVector1.Text
            nrolote = WPartiOri
            WEntra = "N"
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WPartiOri + "'"
            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                With rstLaudo
                    .MoveFirst
                    nrolote = IIf(IsNull(rstLaudo!Laudo), "", Str$(rstLaudo!Laudo))
                    WEntra = "S"
                    rstLaudo.Close
                End With
            End If
                    
            If WEntra = "N" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo.Text + "'"
                ZSql = ZSql + " and Guia.PartiOri = " + "'" + WPartiOri + "'"
                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    With rstMovguia
                        .MoveFirst
                        nrolote = IIf(IsNull(rstMovguia!Lote), "", Str$(rstMovguia!Lote))
                        rstMovguia.Close
                    End With
                End If
            End If
        
        Case "DK", "NS", "NQ"
            WAuxiliar = Left$(Articulo.Text, 3) + "00" + Right$(Articulo.Text, 7)
            WVector1.Col = 3
            WPartiOri = WVector1.Text
            nrolote = WPartiOri
                
            Sql1 = "Select *"
            Sql2 = " FROM EntDev"
            Sql3 = " Where EntDev.Terminado = " + "'" + WAuxiliar + "'"
            Sql4 = " and EntDev.Codigo = " + "'" + nrolote + "'"
            spEntdev = Sql1 + Sql2 + Sql3 + Sql4
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            If rstEntdev.RecordCount > 0 Then
                nrolote = IIf(IsNull(rstEntdev!Lote), "", Str$(rstEntdev!Lote))
                WEntra = "S"
                rstEntdev.Close
            End If
        
        Case Else
            WPartiOri = ""
            WVector1.Col = 7
            nrolote = WVector1.Text
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Laudo.Lote = " + "'" + nrolote + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                rstLaudo.Close
            End If
            
    End Select
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
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
    
    Erase WGrilla
    LugarGrilla = 0
    
    WArticulo = Articulo.Text
    WLote = nrolote

    If WReventa = 1 And Trim(WPartiOri) <> "" Then
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + WPartiOri + "'"
        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        If WLiberada <> 0 Then
                            LugarGrilla = LugarGrilla + 1
                            WGrilla(LugarGrilla, 1) = rstLaudo!Fecha
                            WGrilla(LugarGrilla, 2) = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                            WGrilla(LugarGrilla, 3) = rstLaudo!Laudo
                            WGrilla(LugarGrilla, 4) = Str$(rstLaudo!Liberada)
                            WGrilla(LugarGrilla, 5) = rstLaudo!Orden
                            WGrilla(LugarGrilla, 6) = "Laudo"
                            WGrilla(LugarGrilla, 7) = rstLaudo!Laudo
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstLaudo.Close
        End If
        
            Else
        
        XParam = "'" + WLote + "','" _
                     + WArticulo + "'"
    
        spLaudo = "ListaLaudoArticulo" + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
            Rem If WLiberada <> 0 Then
                LugarGrilla = LugarGrilla + 1
                WGrilla(LugarGrilla, 1) = rstLaudo!Fecha
                WGrilla(LugarGrilla, 2) = Right$(rstLaudo!Fecha, 4) + Mid$(rstLaudo!Fecha, 4, 2) + Left$(rstLaudo!Fecha, 2)
                WGrilla(LugarGrilla, 3) = rstLaudo!Laudo
                WGrilla(LugarGrilla, 4) = Str$(rstLaudo!Liberada)
                WGrilla(LugarGrilla, 5) = rstLaudo!Orden
                WGrilla(LugarGrilla, 6) = "Laudo"
                WGrilla(LugarGrilla, 7) = rstLaudo!Laudo
            Rem End If
            rstLaudo.Close
           
        End If
            
    End If
    
    For Ciclo = 1 To LugarGrilla
    
        WFecha = WGrilla(Ciclo, 1)
        WFechaord = WGrilla(Ciclo, 2)
        WCodigo = WGrilla(Ciclo, 3)
        WCantidad = Val(WGrilla(Ciclo, 4))
        WComprobante = WGrilla(Ciclo, 5)
        WDescri = WGrilla(Ciclo, 6)
        WLote = WGrilla(Ciclo, 7)

        If WDescri = "Guia In" Then
        
            Select Case Val(WComprobante)
                Case 1
                    WObservaciones = "Recepcion de Surfactan"
                Case 2
                    WObservaciones = "Recepcion de Pellital"
                Case 3
                    WObservaciones = "Recepcion de Surfactan II"
                Case 4
                    WObservaciones = "Recepcion de Pellital II"
                Case 5
                    WObservaciones = "Recepcion de Surfactan III"
                Case 6
                    WObservaciones = "Recepcion de Surfactan IV"
                Case 7
                    WObservaciones = "Recepcion de Surfactan V"
                Case 8
                    WObservaciones = "Recepcion de Pellital V"
                Case 9
                    WObservaciones = "Recepcion de Pellital IV"
                Case 10
                    WObservaciones = "Recepcion de Surfactan VI"
                Case 11
                    WObservaciones = "Recepcion de Surfactan VII"
                Case Else
            End Select
            
                Else
                
            spOrden = "ListaOrden" + "'" + WComprobante + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WProveedor = rstOrden!Proveedor
                rstOrden.Close
            End If
        
            WObservaciones = ""
                
            spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WObservaciones = RstProveedor!Nombre
                RstProveedor.Close
            End If
            
        End If
            
        WDesArticulo = ""
            
        spArticulo = "ConsultaArticulo " + " '" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDesArticulo = rstArticulo!Descripcion
            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
            rstArticulo.Close
        End If
                
        With rstFichaMat
            .AddNew
            !Articulo = WArticulo
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WCodigo
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = 0
            !Descripcion = WDesArticulo
            !Observaciones = WObservaciones
            !Lista1 = WDescri
            !Lista2 = ""
            !Lote = WLote
            !Saldo = 0
            !Empresa = NombreEmpresa
            !PartiOri = WPartiOri
            .Update
        End With
    
    Next Ciclo
            
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
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
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                
                If !Tipo = "M" Then
                
                
                    XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                    
                    Rem If Val(XLote(1, 1)) = 0 And rstHoja!Lote <> 0 Then
                    Rem     XLote(1, 1) = rstHoja!Lote
                    Rem     XLote(1, 2) = rstHoja!Cantidad
                    Rem End If
                    
                    For Da = 1 To 3
                        If Val(XLote(Da, 1)) = Val(nrolote) Then
                
                            WArticulo = rstHoja!Articulo
                            WCantidad = XLote(Da, 2)
                            WFecha = rstHoja!Fecha
                            WHoja = rstHoja!Hoja
                            WSaldo = 0
                
                            With rstFichaMat
                
                                .AddNew
                                !Articulo = WArticulo
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WHoja
                                !Inicial = 0
                                !Entrada = 0
                                !Salida = WCantidad
                                !Observaciones = ""
                                !Descripcion = WDesArticulo
                                !Lista1 = "Hoja"
                                !Lista2 = ""
                                !Lote = Val(nrolote)
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                            End With
                        End If
                    Next Da
                        
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
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                
                If !Tipo = "M" Then
                
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    If Val(WLote) = Val(nrolote) Then
                
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        WTipomov = Val(rstMovvar!Tipomov)
                        WObservaciones = rstMovvar!Observaciones
                        WSaldo = 0
                    
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
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
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            If WTipomov = 0 Or WTipomov = 1 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
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
    
    
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
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
                
                If rstMovguia!Marca = "X" Then
                
                        Else
                        
                If rstMovguia!Tipo = "M" Then
            
                    WArticulo = rstMovguia!Articulo
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    Rem WObservaciones = rstMovvar!Observaciones
                        
                    If WMovi = "E" Then
                        WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                            Else
                        WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                        ZPArtiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        WSaldo = 0
                    End If

                        
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
                            
                                Else
                                
                        Select Case WTipomov
                            Case 1
                                WObservaciones = "Recepcion de Surfactan"
                            Case 2
                                WObservaciones = "Recepcion de Pellital"
                            Case 3
                                WObservaciones = "Recepcion de Surfactan II"
                            Case 4
                                WObservaciones = "Recepcion de Pellital II"
                            Case 5
                                WObservaciones = "Recepcion de Surfactan III"
                            Case 6
                                WObservaciones = "Recepcion de Surfactan IV"
                            Case 7
                                WObservaciones = "Recepcion de Surfactan V"
                            Case 8
                                WObservaciones = "Recepcion de Pellital V"
                            Case 9
                                WObservaciones = "Recepcion de Pellital IV"
                            Case 10
                                WObservaciones = "Recepcion de Surfactan VI"
                            Case 11
                                WObservaciones = "Recepcion de Surfactan VII"
                            Case Else
                        End Select
                            
                    End If
                    
                    If WReventa = 1 And Trim(WPartiOri) <> "" Then
                    
                        If Trim(ZPArtiOri) = Trim(WPartiOri) Then
                            With rstFichaMat
                                .AddNew
                                !Articulo = WArticulo
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
                                !Observaciones = WObservaciones
                                !Descripcion = WDesArticulo
                                If !Numero > 900000 Then
                                    !Lista1 = "Prestamo"
                                    !Numero = !Numero - 900000
                                        Else
                                    !Lista1 = "Guia In"
                                End If
                                !Lista2 = ""
                                !Lote = WLote
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                                
                            End With
                        End If
                            
                            Else
                            
                        If WLote = Val(nrolote) Then
                            With rstFichaMat
                    
                                .AddNew
                                !Articulo = WArticulo
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
                                !Observaciones = WObservaciones
                                !Descripcion = WDesArticulo
                                If !Numero > 900000 Then
                                    !Lista1 = "Prestamo"
                                    !Numero = !Numero - 900000
                                        Else
                                    !Lista1 = "Guia In"
                                End If
                                !Lista2 = ""
                                !Lote = WLote
                                !Saldo = WSaldo
                                !Empresa = NombreEmpresa
                                !PartiOri = WPartiOri
                                .Update
                                
                            End With
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
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
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
                
                If !Tipo = "M" Then
                
                    WArticulo = rstMovlab!Articulo
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                    Rem WSaldo = IIf(IsNull(rstMovlab!Saldo), "0", rstMovlab!Saldo)
                    
                    If Val(WLote) = Val(nrolote) Then
                        
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
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
                            !Observaciones = WObservaciones
                            !Descripcion = WDesArticulo
                            !Lista1 = "Mov.Lab"
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !Empresa = NombreEmpresa
                            !PartiOri = WPartiOri
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
    
    Rem PROCESA LAS VENTAS
    
    WTipopro = Left$(Articulo.Text, 2)
    Select Case WTipopro
        Case "DK", "NS", "NQ"
        
            Select Case WTipopro
                Case "DK"
                    ZTipoPro = "DY"
                Case "NS"
                    ZTipoPro = "DS"
                Case Else
                    ZTipoPro = "DQ"
            End Select
            ZZArticulo = ZTipoPro + Mid$(Articulo.Text, 3, 8)
        
            XParam = "'" + ZZArticulo + "','" _
                 + ZZArticulo + "'"
    
            spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
    
                With rstEstadistica
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstEstadistica!Marca = "X" Then
                
                                Else
                
                            If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = ZZArticulo Then
                    
                                If rstEstadistica!Tipo = 2 Then
                
                                    WFecha = rstEstadistica!Fecha
                                    WCodigo = rstEstadistica!Numero
                                    WObservaciones = rstEstadistica!Cliente
                                    WTipo = rstEstadistica!Tipo
                        
                                    WCantidad = rstEstadistica!Canti1
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = WFecha
                                    Vector(Lugar, 2) = "Devol"
                                    Vector(Lugar, 3) = WCodigo
                                    Vector(Lugar, 4) = rstEstadistica!Cliente
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                                    WXSalidas = WXSalidas + WCantidad
                                    Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    Vector(Lugar, 9) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                                    Vector(Lugar, 10) = ""
                                    
                                    WNroLaudo = Vector(Lugar, 9)
                                    WEntra = "N"
                                    WPartiOri = ""
            
                                    ZSql = ""
                                    ZSql = ZSql + "Select *"
                                    ZSql = ZSql + " FROM Laudo"
                                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZArticulo + "'"
                                    ZSql = ZSql + " and Laudo.Laudo = " + "'" + WNroLaudo + "'"
                                    spLaudo = ZSql
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstLaudo.RecordCount > 0 Then
                                        WPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                                        WEntra = "S"
                                        rstLaudo.Close
                                    End If
                        
                                    If WEntra = "N" Then
                                        ZSql = ""
                                        ZSql = ZSql + "Select *"
                                        ZSql = ZSql + " FROM Guia"
                                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZZArticulo + "'"
                                        ZSql = ZSql + " and Guia.Lote = " + "'" + WNroLaudo + "'"
                                        spMovguia = ZSql
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                            WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                                            WEntra = "S"
                                            rstMovguia.Close
                                                Else
                                            WPartiOri = Vector(Cicla, 9)
                                        End If
                                    End If
                                    
                                    With rstFichaMat
                                        .AddNew
                                        !Articulo = WArticulo
                                        !Fecha = WFecha
                                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                        !Tipo = 0
                                        !Numero = WCodigo
                                        !Inicial = 0
                                        !Entrada = 0
                                        !Salida = Abs(Val(WCantidad))
                                        !Lista1 = "Devol."
                                        !Observaciones = WObservaciones
                                        !Descripcion = ""
                                        !Lista2 = ""
                                        !Lote = WLote
                                        !Saldo = 0
                                        !Empresa = NombreEmpresa
                                        !PartiOri = WPartiOri
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
                rstEstadistica.Close
            End If
    
        Case Else
            XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
            spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
    
                With rstEstadistica
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstEstadistica!Marca = "X" Then
                
                                Else
                
                            If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = Articulo.Text Then
                    
                            If (rstEstadistica!Tipo = 2 And Left$(Articulo.Text, 2) = rstEstadistica!Tipopro) Or rstEstadistica!Tipo = 1 Then
     
                                WArticulo = rstEstadistica!ArticuloDy
                                WFecha = rstEstadistica!Fecha
                                WCodigo = rstEstadistica!Numero
                                
                                
                                WObservaciones = rstEstadistica!Cliente
                                WTipo = rstEstadistica!Tipo
                        
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
                                        Else
                                    XLote(6, 1) = "0"
                                    XLote(6, 2) = "0"
                                    XLote(7, 1) = "0"
                                    XLote(7, 2) = "0"
                                    XLote(8, 1) = "0"
                                    XLote(8, 2) = "0"
                                    XLote(9, 1) = "0"
                                    XLote(9, 2) = "0"
                                    XLote(10, 1) = "0"
                                    XLote(10, 2) = "0"
                                    XLote(11, 1) = "0"
                                    XLote(11, 2) = "0"
                                    XLote(12, 1) = "0"
                                    XLote(12, 2) = "0"
                                End If
                        
                                For Da = 1 To 12
                        
                                    WLote = XLote(Da, 1)
                                    WCantidad = XLote(Da, 2)
                                    Rem by nan
                                    
                                      If Val(WLote) = Val(nrolote) Then
                            
                                        With rstFichaMat
                    
                                            .AddNew
                                            !Articulo = WArticulo
                                            !Fecha = WFecha
                                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            !Tipo = 0
                                            !Numero = WCodigo
                                            !Inicial = 0
                                            If WTipo = 2 Then
                                                !Entrada = Abs(Val(WCantidad))
                                                !Salida = 0
                                                !Lista1 = "Devol."
                                                    Else
                                                !Entrada = 0
                                                !Salida = WCantidad
                                                !Lista1 = "Factura"
                                            End If
                                            !Observaciones = WObservaciones
                                            !Descripcion = ""
                                            !Lista2 = ""
                                            !Lote = WLote
                                            !Saldo = 0
                                            !Empresa = NombreEmpresa
                                            !PartiOri = WPartiOri
                                            .Update
                                        End With
                                    End If
                            
                                Next Da
                        
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
                rstEstadistica.Close
            End If
    End Select
    
    Rem PROCESA LOS las devoluciones de mercaderia
    
    WAuxiliar = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
    
    XParam = "'" + WAuxiliar + "','" _
                 + WAuxiliar + "'"
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
                        
                
                    WCantidad = rstEntdev!Cantidad
                    WFecha = rstEntdev!Fecha
                    WCodigo = rstEntdev!Codigo
                    WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                    WPartiOri = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                    WSaldo = rstEntdev!Saldo
                
                    If Val(nrolote) = WLote Then
                        With rstFichaMat
                            .AddNew
                            !Articulo = WArticulo
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            !Entrada = WCantidad
                            !Salida = 0
                            !Observaciones = ""
                            !Lista1 = "Ent.Dev."
                            !Lista2 = ""
                            !Lote = WLote
                            !Saldo = WSaldo
                            !PartiOri = WPartiOri
                            .Update
                        End With
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
    
    
    
    
    
    
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WObservaciones = !Observaciones
                WDescripcion = ""
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                If !Lista1 = "Devol." Or !Lista1 = "Factura" Then
                    spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                End If
                !Descripcion = WDescripcion
                !Observaciones = WObservaciones
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    ZTipo = Left$(WArticulo, 2)
    If WReventa = 1 Then
        ZTipo = "DY"
    End If
    If ZTipo = "DY" Or ZTipo = "DK" Or ZTipo = "DS" Or ZTipo = "NS" Or ZTipo = "DQ" Or ZTipo = "NQ" Then
        Listado.ReportFileName = "WLotematdy.rpt"
            Else
        Listado.ReportFileName = "WLotemat.rpt"
    End If

    Listado.WindowTitle = "Listado de Ficha Lote de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.Destination = 0
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaMat
End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Articulo.Text = rstArticulo!Codigo
        DesArticulo.Caption = rstArticulo!Descripcion
        WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
        rstArticulo.Close
        Call Proceso_Click
        Rem WVector1.SetFocus
            Else
        Articulo.Text = WArticulo
    End If
    Rem Articulo.SetFocus
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsFicMat.Caption = "Consulta de Ficha de Stock de Materias Primas :  " + !Nombre
            NombreEmpresa = !Nombre
        End If
    End With
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        Stock1.Caption = "SI"
        Stock2.Caption = "SII"
        Stock3.Caption = "SIII"
        Stock4.Caption = "SIV"
        Stock5.Caption = "SV"
        Stock6.Caption = "SVI"
        Stock7.Caption = "SVII"
            Else
        Stock1.Caption = "PI"
        Stock2.Caption = "PII"
        Stock3.Caption = "PV"
        Stock4.Caption = "PIV"
        Stock5.Caption = ""
        Stock6.Caption = ""
        Stock7.Caption = ""
    End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Articulo.Text = UCase(Articulo.Text)
    
    WSalidaError = ""
    On Error GoTo Control_Error
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
        
            XEmpresa = WEmpresa
        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock1.Caption = "0"
            End If
        
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock3.Caption = "0"
            End If
    
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock4.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock4.Caption = "0"
            End If
    
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock5.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock5.Caption = "0"
            End If
            
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock6.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock6.Caption = "0"
            End If
            
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock7.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock7.Caption = "0"
            End If
            
            Select Case Val(XEmpresa)
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
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
    
        Case 2, 4, 8, 9
        
            XEmpresa = WEmpresa
    
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                      Else
                WStock1.Caption = "0"
            End If
    
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock2.Caption = "0"
            End If
            
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock3.Caption = "0"
            End If
            
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStock4.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                     Else
                WStock4.Caption = "0"
            End If
            
            Select Case Val(XEmpresa)
                Case 2
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
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
                Case Else
            End Select
    
        Case Else
    End Select
    
    On Error GoTo 0
    
    WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
    WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
    WStock6.Caption = Pusing("###,###.##", WStock6.Caption)
    WStock7.Caption = Pusing("###,###.##", WStock7.Caption)

    WXInicial = 0
    WXEntradas = 0
    WXSalidas = 0
    WXStock = 0

    Renglon = 0
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                
        WArticulo = rstArticulo!Codigo
        WInicial = rstArticulo!Inicial
                                        
        Renglon = Renglon + 1
                   
        WVector1.TextMatrix(Renglon, 1) = WFechaCierre
        WVector1.TextMatrix(Renglon, 2) = ""
        WVector1.TextMatrix(Renglon, 3) = ""
        WVector1.TextMatrix(Renglon, 4) = "Saldo Inicial"
        WVector1.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(rstArticulo!Inicial))
        WVector1.TextMatrix(Renglon, 6) = ""
        WVector1.TextMatrix(Renglon, 7) = ""
        WVector1.TextMatrix(Renglon, 8) = ""
                
        WXInicial = rstArticulo!Inicial
        
        rstArticulo.Close
                
    End If
                
    Rem PROCESA LOS LAUDOS
    
    Erase Vector
    Lugar = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
    Rem ZSql = ZSql + " and Laudo.Laudo = " + "'" + WNroLaudo + "'"
    spLaudo = ZSql
    Rem XParam = "'" + Articulo.Text + "','" _
    rem              + Articulo.Text + "'"
    Rem spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
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
                    
                    If rstLaudo!Articulo = Articulo.Text Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!Liberada
                        WFecha = rstLaudo!Fecha
                        WLaudo = rstLaudo!Laudo
                        WPartiOri = rstLaudo!PartiOri
                        WOrden = rstLaudo!Orden
                        WDevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        WRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        Call Redondeo(WSaldo)
                        
                        Rem If WLiberada <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Laudo"
                            Vector(Lugar, 3) = WLaudo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(WLiberada))
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WOrden
                            If WReventa = 1 Then
                                Vector(Lugar, 9) = Left$(WPartiOri, 10)
                                    Else
                                Vector(Lugar, 9) = WLaudo
                            End If
                            Vector(Lugar, 10) = Str$(WSaldo)
                    
                            WXEntradas = WXEntradas + WLiberada
                            
                        Rem End If
                        
                        If WDevuelta <> 0 Then
                
                            Lugar = Lugar + 1
                        
                            Vector(Lugar, 1) = !Fecha
                            Vector(Lugar, 2) = "Rechazo"
                            Vector(Lugar, 3) = WRechazo
                            Vector(Lugar, 4) = WDEsProveedor
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = "(" + Pusing("###,###.##", Str$(rstLaudo!devuelta)) + ")"
                            Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            Vector(Lugar, 8) = WOrden
                            Vector(Lugar, 9) = WRechazo
                            Vector(Lugar, 10) = "0"
                            
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
    
    For Ciclo = 1 To Lugar
    
        WOrden = Vector(Ciclo, 8)
        
        WProveedor = ""
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WDEsProveedor = RstProveedor!Nombre
            RstProveedor.Close
        End If
    
        Vector(Ciclo, 4) = WDEsProveedor
        
    Next Ciclo
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        
        WDesvio = "N"
        If Val(Vector(Cicla, 9)) >= 190000 And Val(Vector(Cicla, 9)) <= 194999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 990000 And Val(Vector(Cicla, 9)) <= 994999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 290000 And Val(Vector(Cicla, 9)) <= 294999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 390000 And Val(Vector(Cicla, 9)) <= 394999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 490000 And Val(Vector(Cicla, 9)) <= 494999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 590000 And Val(Vector(Cicla, 9)) <= 594999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 690000 And Val(Vector(Cicla, 9)) <= 694999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 790000 And Val(Vector(Cicla, 9)) <= 794999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 890000 And Val(Vector(Cicla, 9)) <= 894999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                                               
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
        
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
    
    Next Cicla
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
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
                Rem If XFec < WOrdFechaCierre Then
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                        
                    fr = rstHoja!Clave
                        
                    If rstHoja!Tipo = "M" And rstHoja!Articulo = Articulo.Text Then
                    
                
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) = "" Then
                                XLote(Da, 2) = "0"
                            End If
                        
                            WCanti = XLote(Da, 2)
                            If WCanti <> 0 Then
                
                                WArticulo = rstHoja!Articulo
                                WCanti = XLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = XLote(Da, 1)
                        
                                Lugar = Lugar + 1
                        
                                Vector(Lugar, 1) = !Fecha
                                Vector(Lugar, 2) = "Hoja"
                                Vector(Lugar, 3) = WHoja
                                Vector(Lugar, 4) = ""
                                Vector(Lugar, 5) = ""
                                Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCanti * 1))
                                Vector(Lugar, 7) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                Vector(Lugar, 9) = WLote
                                Vector(Lugar, 10) = ""
                        
                                WXSalidas = WXSalidas + WCanti
                                
                            End If
                        Next Da

                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If !Articulo > Articulo.Text Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
    
    Next Cicla
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
    spMovvar = "ListaMovvarArticuloDesdeHasta" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                        
                    If rstMovvar!Tipo = "M" And rstMovvar!Articulo = Articulo.Text Then
                    
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovvar!Fecha
                        If rstMovvar!Tipomov = 0 Or rstMovvar!Tipomov = 1 Then
                            Vector(Lugar, 2) = "Mov.Var"
                                Else
                            Vector(Lugar, 2) = "Guia In"
                        End If
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovvar!Observaciones
                        If rstMovvar!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovvar!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovvar!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(rstMovvar!Cantidad))
                            WXSalidas = WXSalidas + rstMovvar!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        Vector(Lugar, 10) = ""
                        
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
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
    
    Next Cicla
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                + Articulo.Text + "'"
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
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = Articulo.Text Then
                    
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovguia!Fecha
                        If Val(WCodigo) > 900000 Then
                            Vector(Lugar, 2) = "Prestamo"
                            Vector(Lugar, 3) = WCodigo - 900000
                                Else
                            Vector(Lugar, 2) = "Guia In"
                            Vector(Lugar, 3) = WCodigo
                        End If
                        Rem Vector(Lugar, 4) = rstMovguia!Observaciones
                                
                        If rstMovguia!Movi = "E" Then
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
                                    If Left$(Articulo.Text, 2) = "DW" Then
                                        Vector(Lugar, 2) = "Hoja"
                                        Vector(Lugar, 3) = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                                        Vector(Lugar, 4) = "Hoja de Produccion"
                                            Else
                                        Vector(Lugar, 4) = "Recepcion de Pellital IV"
                                    End If
                                Case 10
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovguia!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovguia!Cantidad
                            WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            If Trim(WPartiOri) <> "" Then
                                Vector(Lugar, 9) = WPartiOri
                                    Else
                                Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            Call Redondeo(WSaldo)
                            Vector(Lugar, 10) = Str$(WSaldo)
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
                                    Vector(Lugar, 4) = "Envio a Pellital IV"
                                Case 10
                                    Vector(Lugar, 4) = "Envio a Surfactan VI"
                                Case 11
                                    Vector(Lugar, 4) = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(rstMovguia!Cantidad))
                            WXSalidas = WXSalidas + rstMovguia!Cantidad
                            WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                            If Trim(WPartiOri) <> "" Then
                                Vector(Lugar, 9) = WPartiOri
                                    Else
                                Vector(Lugar, 9) = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            End If
                            Vector(Lugar, 10) = ""
                        End If
                        Vector(Lugar, 7) = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                        
                        
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

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
        
        WDesvio = "N"
        If Val(Vector(Cicla, 9)) >= 190000 And Val(Vector(Cicla, 9)) <= 194999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 990000 And Val(Vector(Cicla, 9)) <= 994999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 290000 And Val(Vector(Cicla, 9)) <= 294999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 390000 And Val(Vector(Cicla, 9)) <= 394999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 490000 And Val(Vector(Cicla, 9)) <= 494999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 590000 And Val(Vector(Cicla, 9)) <= 594999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 690000 And Val(Vector(Cicla, 9)) <= 694999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 790000 And Val(Vector(Cicla, 9)) <= 794999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
        If Val(Vector(Cicla, 9)) >= 890000 And Val(Vector(Cicla, 9)) <= 894999 Then
            If Val(Vector(Cicla, 10)) <> 0 Then
                WDesvio = "S"
            End If
        End If
                
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                                               
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                        
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
                
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
    
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
        If WDesvio = "S" Then
            WVector1.CellBackColor = &H8080FF
        End If
    
    Next Cicla
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHasta" + XParam
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
                
                    If rstMovlab!Tipo = "M" And rstMovlab!Articulo = Articulo.Text Then
                
                        WArticulo = rstMovlab!Articulo
                        WCantidad = rstMovlab!Cantidad
                        WFecha = rstMovlab!Fecha
                        WCodigo = rstMovlab!Codigo
                        WMovi = rstMovlab!Movi
                        
                        Lugar = Lugar + 1
                        
                        Vector(Lugar, 1) = rstMovlab!Fecha
                        Vector(Lugar, 2) = "Mov.Lab"
                        Vector(Lugar, 3) = WCodigo
                        Vector(Lugar, 4) = rstMovlab!Observaciones
                        If rstMovlab!Movi = "E" Then
                            Vector(Lugar, 5) = Pusing("###,###.##", Str$(rstMovlab!Cantidad))
                            Vector(Lugar, 6) = ""
                            WXEntradas = WXEntradas + rstMovlab!Cantidad
                                Else
                            Vector(Lugar, 5) = ""
                            Vector(Lugar, 6) = Pusing("###,###.##", Str$(rstMovlab!Cantidad))
                            WXSalidas = WXSalidas + rstMovlab!Cantidad
                        End If
                        Vector(Lugar, 7) = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                        Vector(Lugar, 9) = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                        Vector(Lugar, 10) = ""
                        
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
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
    
    Next Cicla
    
    Rem PROCESA LAS VENTAS
    
    Erase Vector
    Lugar = 0
    
    WTipopro = Left$(Articulo.Text, 2)
    Select Case WTipopro
        Case "DK", "NS", "NQ"
        
            Select Case WTipopro
                Case "DK"
                    ZTipoPro = "DY"
                Case "NS"
                    ZTipoPro = "DS"
                Case Else
                    ZTipoPro = "DQ"
            End Select
            ZZArticulo = ZTipoPro + Mid$(Articulo.Text, 3, 8)
        
            XParam = "'" + ZZArticulo + "','" _
                 + ZZArticulo + "'"
    
            spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
    
                With rstEstadistica
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstEstadistica!Marca = "X" Then
                
                                Else
                
                            If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = ZZArticulo Then
                    
                                If rstEstadistica!Tipo = 2 Then
                
                                    WArticulo = rstEstadistica!ArticuloDy
                                    WFecha = rstEstadistica!Fecha
                                    WCodigo = rstEstadistica!Numero
                        
                                    WCantidad = rstEstadistica!Canti1
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = WFecha
                                    Vector(Lugar, 2) = "Devol"
                                    Vector(Lugar, 3) = WCodigo
                                    Vector(Lugar, 4) = rstEstadistica!Cliente
                                    Vector(Lugar, 5) = ""
                                    Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                                    WXSalidas = WXSalidas + WCantidad
                                    Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    Vector(Lugar, 9) = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                                    Vector(Lugar, 10) = ""
                        
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
        
        Case Else
            XParam = "'" + Articulo.Text + "','" _
                 + Articulo.Text + "'"
    
            spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
    
                With rstEstadistica
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstEstadistica!Marca = "X" Then
                
                                Else
                
                            If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = Articulo.Text Then
                    
                                If (rstEstadistica!Tipo = 2 And Left$(Articulo.Text, 2) = rstEstadistica!Tipopro) Or rstEstadistica!Tipo = 1 Then
                
                                    WArticulo = rstEstadistica!ArticuloDy
                                    WFecha = rstEstadistica!Fecha
                                    WCodigo = rstEstadistica!Numero
                        
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
                                            Else
                                        XLote(6, 1) = "0"
                                        XLote(6, 2) = "0"
                                        XLote(7, 1) = "0"
                                        XLote(7, 2) = "0"
                                        XLote(8, 1) = "0"
                                        XLote(8, 2) = "0"
                                        XLote(9, 1) = "0"
                                        XLote(9, 2) = "0"
                                        XLote(10, 1) = "0"
                                        XLote(10, 2) = "0"
                                        XLote(11, 1) = "0"
                                        XLote(11, 2) = "0"
                                        XLote(12, 1) = "0"
                                        XLote(12, 2) = "0"
                                    End If
                        
                                    For Da = 1 To 12
                            
                                        WLote = XLote(Da, 1)
                                        Auxi = XLote(Da, 2)
                                        Auxi = Pusing("###,###.##", Auxi)
                                        WCanti = Val(Auxi)
                        
                                        If WCanti <> 0 Then
                                            WCantidad = WCanti
                                            Lugar = Lugar + 1
                                            Vector(Lugar, 1) = WFecha
                                            If rstEstadistica!Tipo = 1 Then
                                                Vector(Lugar, 2) = "Factura"
                                                    Else
                                                WTipopro = rstEstadistica!Tipopro
                                                Vector(Lugar, 2) = "Devol"
                                            End If
                                            Vector(Lugar, 3) = WCodigo
                                            Vector(Lugar, 4) = rstEstadistica!Cliente
                                            If rstEstadistica!Tipo = 2 Then
                                                Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                                                Vector(Lugar, 6) = ""
                                                WXEntradas = WXEntradas + WCantidad
                                                    Else
                                                Vector(Lugar, 5) = ""
                                                Vector(Lugar, 6) = Pusing("###,###.##", Str$(WCantidad))
                                                WXSalidas = WXSalidas + WCantidad
                                            End If
                                            Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                            Vector(Lugar, 9) = WLote
                                            Vector(Lugar, 10) = ""
                                        End If
                        
                                    Next Da
                        
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
            
    End Select
    
    For Ciclo = 1 To Lugar

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        
        spCliente = "ConsultaCliente" + "'" + Vector(Cicla, 4) + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WVector1.TextMatrix(Renglon, 4) = rstCliente!Razon
                Else
            WVector1.TextMatrix(Renglon, 4) = ""
        End If
                        
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        
        If WReventa = 1 Then
            WNroLaudo = Vector(Cicla, 9)
            WEntra = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo.Text + "'"
            ZSql = ZSql + " and Laudo.Laudo = " + "'" + WNroLaudo + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WVector1.TextMatrix(Renglon, 7) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                WEntra = "S"
                rstLaudo.Close
            End If
                        
            If WEntra = "N" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo.Text + "'"
                ZSql = ZSql + " and Guia.Lote = " + "'" + WNroLaudo + "'"
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WVector1.TextMatrix(Renglon, 7) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                    WEntra = "S"
                    rstMovguia.Close
                        Else
                    WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
                End If
                
            End If
            
                Else
                
            WNroLaudo = Vector(Cicla, 9)
            WEntra = "N"
            WPartiOri = ""
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZArticulo + "'"
            ZSql = ZSql + " and Laudo.Laudo = " + "'" + WNroLaudo + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                WEntra = "S"
                rstLaudo.Close
            End If
                        
            If WEntra = "N" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.Articulo = " + "'" + ZZArticulo + "'"
                ZSql = ZSql + " and Guia.Lote = " + "'" + WNroLaudo + "'"
                spMovguia = ZSql
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                    WEntra = "S"
                    rstMovguia.Close
                        Else
                    WPartiOri = Vector(Cicla, 9)
                End If
                
            End If
                
            WVector1.TextMatrix(Renglon, 7) = WPartiOri
            
        End If
        
        WSaldo = Val(Vector(Cicla, 10))
        Call Redondeo(WSaldo)
        WVector1.TextMatrix(Renglon, 8) = Str$(WSaldo)
    
    Next Cicla
    
    
        
        
        
        
        
    Rem PROCESA LOS las devoluciones de mercaderia
    
    WAuxiliar = Left$(Articulo.Text, 3) + "00" + Right$(Articulo.Text, 7)
    
    Erase Vector
    Lugar = 0
    
    XParam = "'" + WAuxiliar + "','" _
                 + WAuxiliar + "'"
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
                
                    WTerminado = rstEntdev!Terminado
                    WCantidad = rstEntdev!Cantidad
                    WLaboratorio = rstEntdev!Laboratorio
                    WFecha = rstEntdev!Fecha
                    WCodigo = rstEntdev!Codigo
                    WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                    WPartiOri = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                    WSaldo = rstEntdev!Saldo
                    
                    Lugar = Lugar + 1
                    
                    Vector(Lugar, 1) = WFecha
                    Vector(Lugar, 2) = "Ent.Dev"
                    Vector(Lugar, 3) = WCodigo
                    Vector(Lugar, 4) = rstEntdev!Observaciones
                    Vector(Lugar, 5) = Pusing("###,###.##", Str$(WCantidad))
                    Vector(Lugar, 6) = ""
                    WXEntradas = WXEntradas + WCantidad
                    Vector(Lugar, 7) = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    Vector(Lugar, 8) = WLote
                    Vector(Lugar, 9) = WPartiOri
                    Vector(Lugar, 10) = Pusing("###,###.##", Str$(WSaldo))
                
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

        For dada = Ciclo + 1 To Lugar

            If Vector(Ciclo, 7) > Vector(dada, 7) Then

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
                
                Vector(Ciclo, 1) = Vector(dada, 1)
                Vector(Ciclo, 2) = Vector(dada, 2)
                Vector(Ciclo, 3) = Vector(dada, 3)
                Vector(Ciclo, 4) = Vector(dada, 4)
                Vector(Ciclo, 5) = Vector(dada, 5)
                Vector(Ciclo, 6) = Vector(dada, 6)
                Vector(Ciclo, 7) = Vector(dada, 7)
                Vector(Ciclo, 8) = Vector(dada, 8)
                Vector(Ciclo, 9) = Vector(dada, 9)
                Vector(Ciclo, 10) = Vector(dada, 10)
                
                Vector(dada, 1) = Auxi1
                Vector(dada, 2) = Auxi2
                Vector(dada, 3) = Auxi3
                Vector(dada, 4) = Auxi4
                Vector(dada, 5) = Auxi5
                Vector(dada, 6) = Auxi6
                Vector(dada, 7) = Auxi7
                Vector(dada, 8) = Auxi8
                Vector(dada, 9) = Auxi9
                Vector(dada, 10) = Auxi10

            End If

        Next dada

    Next Ciclo
    
    For Cicla = 1 To Lugar
    
        Renglon = Renglon + 1
                
        WVector1.TextMatrix(Renglon, 1) = Vector(Cicla, 1)
        WVector1.TextMatrix(Renglon, 2) = Vector(Cicla, 2)
        WVector1.TextMatrix(Renglon, 3) = Vector(Cicla, 3)
        WVector1.TextMatrix(Renglon, 4) = Vector(Cicla, 4)
        WVector1.TextMatrix(Renglon, 5) = Vector(Cicla, 5)
        WVector1.TextMatrix(Renglon, 6) = Vector(Cicla, 6)
        WVector1.TextMatrix(Renglon, 7) = Vector(Cicla, 9)
        WVector1.TextMatrix(Renglon, 8) = Vector(Cicla, 10)
    
    Next Cicla

    
   
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    XEmpresa = WEmpresa
    ZSumaSolic = 0
    
    For Cicla = 1 To 11
    
        Select Case Cicla
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
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Solic"
        ZSql = ZSql + " Where Solic.Marca <> " + "'" + "X" + "'"
        ZSql = ZSql + " and Solic.Articulo = " + "'" + Articulo.Text + "'"
        spSolic = ZSql
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
            With rstSolic
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        ZSuma = rstSolic!Cantidad - rstSolic!Entregado
                        ZSumaSolic = ZSumaSolic + ZSuma
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
        
            End With
            rstSolic.Close
        End If
        
    Next Cicla
    
    Call Conecta_Empresa
    
    
    
    
    
    WXStock = WXInicial + WXEntradas - WXSalidas
    
    XInicial.Text = Pusing("###,###.##", Str$(WXInicial))
    XEntradas.Text = Pusing("###,###.##", Str$(WXEntradas))
    XSalidas.Text = Pusing("###,###.##", Str$(WXSalidas))
    XStock.Text = Pusing("###,###.##", Str$(WXStock))
    Solicitud.Text = Pusing("###,###.##", Str$(ZSumaSolic))

    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
    PrgConsFicMat.WindowState = 0
    
    Exit Sub
    
Control_Error:
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    StockCons.Visible = False
    Resume Next
    
End Sub


Private Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        WArticulo = Articulo.Text
        Articulo.Text = WArticulo
        
        spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
            rstArticulo.Close
            Call Proceso_Click
            Rem WVector1.SetFocus
                Else
            Articulo.SetFocus
        End If
        
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 9
    WVector1.FixedRows = 1
    WVector1.Rows = 5001
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WVector1.Text = "Entradas"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Salidas"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 1100
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
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub



