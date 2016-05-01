VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedCentroPelli 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Pedidos"
   ClientHeight    =   8415
   ClientLeft      =   90
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   11850
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Muestra4 
      Height          =   1695
      Left            =   1440
      TabIndex        =   52
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   100
      Cols            =   9
      BackColor       =   16777088
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra3 
      Height          =   1695
      Left            =   960
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   100
      Cols            =   8
      BackColor       =   16777088
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra1 
      Height          =   1215
      Left            =   0
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   100
      Cols            =   8
      BackColor       =   16777088
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra2 
      Height          =   1695
      Left            =   360
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   100
      Cols            =   8
      BackColor       =   16777088
   End
   Begin VB.Frame AuxiliarIngresoIII 
      Height          =   1335
      Left            =   1560
      TabIndex        =   43
      Top             =   3240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton ImpreMp 
         Caption         =   "Impre M.P."
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
         Left            =   5640
         TabIndex        =   53
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton VerMp 
         Caption         =   "Ver M.P."
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
         Left            =   4080
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox AuxiIngresoVII 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   49
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox AuxiIngresoVI 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   44
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   315
         Left            =   480
         TabIndex        =   50
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   315
         Left            =   480
         TabIndex        =   45
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame AuxiliarIngresoII 
      Height          =   1335
      Left            =   1560
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox AuxiIngresoV 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   42
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox AuxiIngresoIV 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   39
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   315
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Partida"
         Height          =   315
         Left            =   480
         TabIndex        =   40
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Orden 
      Height          =   1215
      Left            =   1920
      TabIndex        =   37
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   100
      BackColor       =   16777088
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
      Height          =   615
      Left            =   10320
      TabIndex        =   34
      Top             =   2160
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   3495
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   100
      Cols            =   8
   End
   Begin VB.ComboBox Tipoped 
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
      Left            =   8280
      TabIndex        =   26
      Top             =   480
      Width           =   1575
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.Frame Datos 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3495
      Left            =   3840
      TabIndex        =   14
      Top             =   4800
      Width           =   3255
      Begin VB.TextBox Termi 
         BackColor       =   &H00FFFF00&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "Produccion"
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
         TabIndex        =   47
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label StkProduccion 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   46
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Descripcion 
         BackColor       =   &H00FFFF00&
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
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   3015
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
         Left            =   1920
         TabIndex        =   33
         Top             =   2880
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
         Left            =   120
         TabIndex        =   32
         Top             =   2880
         Width           =   1815
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
         Left            =   1920
         TabIndex        =   31
         Top             =   2640
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
         TabIndex        =   30
         Top             =   2640
         Width           =   1695
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
         TabIndex        =   24
         Top             =   2400
         Width           =   1695
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
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
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
         Left            =   1920
         TabIndex        =   22
         Top             =   2400
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
         Left            =   1920
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Disponible 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label StkPedido 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Stock 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Stock Disponible"
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
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Pedido Pendiente"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1695
      End
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
      Height          =   615
      Left            =   10320
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
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
      Left            =   4920
      MaxLength       =   50
      TabIndex        =   12
      Text            =   " "
      Top             =   840
      Width           =   6615
   End
   Begin MSMask.MaskEdBox FecEntrega 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Cliente 
      BeginProperty Font 
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
      TabIndex        =   6
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Pedido 
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
      TabIndex        =   2
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo Pedido"
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
      TabIndex        =   25
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
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
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Entrega"
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
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesPago 
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
      Left            =   7560
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label DesCliente 
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
      Left            =   3120
      TabIndex        =   7
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cliente"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numero de pedido"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgPedCentroPelli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Lugar3 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WImpre(10) As String
Private WVector(6, 3) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private Transfe(100, 5) As String
Private XSaldo As Double
Dim rstSolGuia As Recordset
Dim spSolGuia As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstImpreMP As Recordset
Dim spImpreMP As String

Dim XParam As String
Dim ClavePedido(100)
Dim WProceso As Integer
Dim WSaldo As Double
Dim WNumeroSolGuia As String
Dim XPedido As Double
Dim XProdu As Double

Dim ZZStockPt As Double
Dim ZZStockMp As Double
Dim ZZTerminado As String
Dim ZZArticulo As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgPedCentroPelli.Hide
    Unload Me
    PrgCentro.Show
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()

    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 1600
    Muestra.ColWidth(2) = 3000
    Muestra.ColWidth(3) = 900
    Muestra.ColWidth(4) = 900
    Muestra.ColWidth(5) = 1000
    Muestra.ColWidth(6) = 1000
    Muestra.ColWidth(7) = 3000
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Producto"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 4
    Muestra.Text = "Saldo"
    
    Muestra.Col = 5
    Muestra.Text = "Proceso"
    
    Muestra.Col = 6
    Muestra.Text = "Part./Cant."
    
    Muestra.Col = 7
    Muestra.Text = "Observ."
  
    Tipoped.Clear
    
    Tipoped.AddItem "Normal"
    Tipoped.AddItem "a Fecha"
    Tipoped.AddItem "Fecha Limite"
    Tipoped.AddItem "Urgente"
    Tipoped.AddItem "Retira Cliente"
    Tipoped.AddItem "Muestra"
    
    Tipoped.ListIndex = 0
    
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

    Pedido.Text = WXPed
    
    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
            Fecha.Text = rstPedido!Fecha
            Cliente.Text = rstPedido!Cliente
            FecEntrega.Text = rstPedido!FecEntrega
            Observaciones.Text = rstPedido!Observaciones
            Tipoped.ListIndex = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
            rstPedido.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                Observaciones.Text = rstCliente!Observaciones
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Pedido.Text = WPedido
    End If
    
    Call Conecta_Empresa
    
End Sub

Private Sub Proceso_Click()

    Muestra.Clear
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Producto"
    
    Muestra.Col = 2
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 3
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 4
    Muestra.Text = "Saldo"
    
    Muestra.Col = 5
    Muestra.Text = "Estado"
    
    Muestra.Col = 6
    Muestra.Text = "Part./Cant."
    
    Muestra.Col = 7
    Muestra.Text = "Observ."
    
    Erase Auxiliar
    Erase Transfe
    Erase ClavePedido
    
    Renglon = 0
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        Renglon = Renglon + 1
                        
                        Muestra.Row = Renglon
                
                        Muestra.Col = 1
                        Muestra.Text = rstPedido!Terminado
                        Auxi1 = rstPedido!Terminado
                
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", rstPedido!Cantidad)
                
                        Muestra.Col = 4
                        Muestra.Text = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!Terminado
                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                            Auxiliar(WRenglon, 3) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
                        End If
                        
                        ClavePedido(WRenglon) = rstPedido!Clave
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
    End If
    
    Renglon = 0
    Total = 0
    
    For Da = 1 To WRenglon
    
        Cliente = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        ZZNombreComercial = Trim(Auxiliar(Da, 3))
        
        If Left$(Terminado, 2) = "PT" Or Left$(Terminado, 2) = "YQ" Or Left$(Terminado, 2) = "YF" Or Left$(Terminado, 2) = "YP" Then
            WTipopro = "T"
                Else
            WTipopro = "M"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Terminado, 3) + Right$(Terminado, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    Muestra.Row = Renglon
                    
                    If ZZNombreComercial <> "" Then
                        Muestra.Col = 2
                        Muestra.Text = ZZNombreComercial
                            Else
                        Muestra.Col = 2
                        Muestra.Text = rstArticulo!Descripcion
                    End If
                    
                    rstArticulo.Close
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Renglon = Renglon + 1
                    Muestra.Row = Renglon
                    Muestra.Col = 2
                    Muestra.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
        End Select
        
    Next Da
    
    Muestra.Row = 1

End Sub

Private Sub ImpreMp_Click()

    Dim ZZImpreMp(100, 10) As String
    
    Erase ZZImpreMp
    Renglon = 0
    ZZCantidad = 0
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Termi.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
        
                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
        
                        Renglon = Renglon + 1
                        
                        ZZImpreMp(Renglon, 1) = rstComposicion!Tipo
                        ZZImpreMp(Renglon, 2) = rstComposicion!Articulo2
                        Auxi2 = rstComposicion!Articulo2
                    
                        If rstComposicion!Articulo1 = "  -   -  " Then
                            ZZImpreMp(Renglon, 3) = "  -   -   "
                            Auxi1 = "  -   -   "
                                Else
                            ZZImpreMp(Renglon, 3) = rstComposicion!Articulo1
                            Auxi1 = rstComposicion!Articulo1
                        End If
                    
                        Cantidad = Str$(rstComposicion!Cantidad * Val(AuxiIngresoVII))
                        ZZCantidad = rstComposicion!Cantidad
                        
                        Auxi2 = Cantidad
                        Auxi2 = Pusing("###,###.##", Auxi2)
                        ZZImpreMp(Renglon, 5) = Auxi2
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    
    ZSql = "DELETE ImpreMP"
    spImpreMP = ZSql
    Set rstImpreMP = db.OpenRecordset(spImpreMP, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZDesTermi = ""
    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZZDesTermi = rstTerminado!Descripcion
        rstTerminado.Close
    End If
    
    For ZZCiclo = 1 To Renglon
    
        ZZTipo = ZZImpreMp(ZZCiclo, 1)
        ZZTerminado = ZZImpreMp(ZZCiclo, 2)
        ZZArticulo = ZZImpreMp(ZZCiclo, 3)
    
        Select Case ZZTipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + ZZTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZImpreMp(ZZCiclo, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZImpreMp(ZZCiclo, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                XEmpresa = WEmpresa
        
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZStockII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    ZStockII = 0
                End If
                Auxi2 = Str$(ZStockII)
                Auxi2 = Pusing("###,###.##", Auxi2)
                ZZImpreMp(ZZCiclo, 6) = Auxi2
                
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                      
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZStockIII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    ZStockIII = 0
                End If
                Auxi2 = Str$(ZStockIII)
                Auxi2 = Pusing("###,###.##", Auxi2)
                ZZImpreMp(ZZCiclo, 7) = Auxi2
    
                Auxi2 = ZStockII + ZStockIII - Val(ZZImpreMp(ZZCiclo, 5))
                Auxi2 = Pusing("###,###.##", Auxi2)
                ZZImpreMp(ZZCiclo, 8) = Auxi2
    
                Call Conecta_Empresa
                
            Case Else
        End Select
        
        Auxi = Str$(ZZCiclo)
        Call Ceros(Auxi, 2)
        ZZClave = Termi.Text + Auxi
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreMP ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "DescripcionII ,"
        ZSql = ZSql + "CantidadII ,"
        ZSql = ZSql + "StockII ,"
        ZSql = ZSql + "StockIV ,"
        ZSql = ZSql + "Diferencia )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + Termi.Text + "',"
        ZSql = ZSql + "'" + Str$(ZZCiclo) + "',"
        ZSql = ZSql + "'" + ZZDesTermi + "',"
        ZSql = ZSql + "'" + AuxiIngresoVII + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 1) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 2) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 3) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 4) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 5) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 6) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 7) + "',"
        ZSql = ZSql + "'" + ZZImpreMp(ZZCiclo, 8) + "')"
        
        spImpreMP = ZSql
        Set rstImpreMP = db.OpenRecordset(spImpreMP, dbOpenSnapshot, dbSQLPassThrough)
    
    Next ZZCiclo
    
    Listado.WindowTitle = ""
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ImpreMp.Clave, ImpreMp.Codigo, ImpreMp.Renglon, ImpreMp.Descripcion, ImpreMp.Cantidad, ImpreMp.Tipo, ImpreMp.Terminado, ImpreMp.Articulo, ImpreMp.DescripcionII, ImpreMp.CantidadII, ImpreMp.StockII, ImpreMp.StockIV, ImpreMp.Diferencia " _
            + "From " _
            + DSQ + ".dbo.ImpreMp ImpreMp " _
            + "Where " _
            + "ImpreMp.Clave >= '' AND " _
            + "ImpreMp.Clave <= 'ZZZZZZZZ'"
            
    Listado.ReportFileName = "ImpreMp.rpt"
    Listado.Connect = Connect()
    Rem Listado.GroupSelectionFormula = "{HistoriaTerminado.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.Action = 1
    
End Sub

Private Sub Muestra_Click()
    Call Muestra_DblClick
End Sub

Private Sub Muestra_DblClick()

    If Muestra.Col = 1 Then

        Muestra.Col = 1
        Termi.Text = Muestra.Text
        XProducto = Termi.Text
        
        Muestra.Col = 2
        Descripcion.Caption = Muestra.Text
        
        Renglon = 0
        XStock = 0
        XPedido = 0
        XProdu = 0
        
        If Left$(XProducto, 2) = "PT" Or Left$(XProducto, 2) = "YQ" Or Left$(XProducto, 2) = "YF" Or Left$(XProducto, 2) = "YP" Then
            WTipopro = "T"
                Else
            WTipopro = "M"
        End If
    
        Call Stock_Consolidado
        
        XStock = (Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption))
        Stock.Caption = Pusing("###,###.##", Str$(XStock))
        aa = Muestra.TextMatrix(Muestra.Row, 4)
        StkPedido.Caption = Pusing("###,###.##", Str$(XPedido - Val(Muestra.TextMatrix(Muestra.Row, 4))))
        StkProduccion.Caption = Pusing("###,###.##", Str$(XProdu))
        Disponible.Caption = Pusing("###,###.##", Str$(XStock - Val(StkPedido.Caption) + XProdu))
    
    End If
    
    If Muestra.Col = 5 Then
        Lugar1 = Muestra.TopRow
        Lugar2 = Muestra.Row
        Lugar3 = Muestra.Col
        Call AyudaOrden
    End If
    
End Sub

Private Sub Stock_Consolidado()

    Termi.Text = UCase(Termi.Text)
    
    Stock1.Caption = "Pta I"
    Stock2.Caption = "Pta II"
    Stock3.Caption = "Pta V"
    Stock4.Caption = "Pta IV"
    
    If Left$(Termi.Text, 2) = "PT" Or Left$(Termi.Text, 2) = "YQ" Or Left$(Termi.Text, 2) = "YF" Or Left$(Termi.Text, 2) = "YP" Then
        WTipopro = "T"
            Else
        WTipopro = "M"
    End If
    WArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
    
    XEmpresa = WEmpresa
    
    WEmpresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
        Rem XPedido = IIf(IsNull(rstTerminado!Pedido), "0", rstTerminado!Pedido)
        rstTerminado.Close
            Else
        WStock1.Caption = "0"
    End If
        
    WEmpresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
        rstTerminado.Close
            Else
        WStock2.Caption = "0"
    End If
                
    WEmpresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
          
    spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
        rstTerminado.Close
            Else
        WStock3.Caption = "0"
    End If
            
    WEmpresa = "0009"
    txtOdbc = "Empresa09"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
    spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
        rstTerminado.Close
            Else
        WStock4.Caption = "0"
    End If
    
    Call Conecta_Empresa
            
    WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
    
    XPedido = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Terminado = " + "'" + Termi.Text + "'"
    ZSql = ZSql + " and Pedido.Cantidad > Pedido.Facturado"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    Rem If Val(Pedido.Text) <> rstPedido!Pedido Then
                        ZZCanti = rstPedido!Cantidad - rstPedido!Facturado
                        XPedido = XPedido + ZZCanti
                    Rem End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If

    
    XProdu = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSolicitud"
    ZSql = ZSql + " Where CargaSolicitud.Articulo = " + "'" + Termi.Text + "'"
    ZSql = ZSql + " and CargaSolicitud.Saldo > 0"
    spCargaSolicitud = ZSql
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSolicitud.RecordCount > 0 Then
        With rstCargaSolicitud
            .MoveFirst
            Do
                If .EOF = False Then
                    XProdu = XProdu + rstCargaSolicitud!Saldo
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaSolicitud.Close
    End If

End Sub

Private Sub Stock_ConsolidadoMP()
    
    XEmpresa = WEmpresa
    
    WEmpresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        ZStockI = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
        rstArticulo.Close
            Else
        ZStockI = 0
    End If
        
    WEmpresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        ZStockII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
        rstArticulo.Close
            Else
        ZStockII = 0
    End If
                
    WEmpresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
          
    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        ZStockIII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
        rstArticulo.Close
            Else
        ZStockIII = 0
    End If
            
    WEmpresa = "0009"
    txtOdbc = "Empresa09"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
    spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        ZStockIV = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
        rstArticulo.Close
            Else
        ZStockIV = 0
    End If
    
    Call Conecta_Empresa
    
    ZZStockMp = ZStockI + ZStockII + ZStockIII + ZStockIV

End Sub


Private Sub Muestra_Ficha()

    XEmpresa = WEmpresa
    XProducto = Termi.Text
    
    Select Case WProceso
        Case 1
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    Muestra1.Clear
    
    Muestra1.ColWidth(0) = 150
    Muestra1.ColWidth(1) = 1300
    Muestra1.ColWidth(2) = 1200
    Muestra1.ColWidth(3) = 1200
    Muestra1.ColWidth(4) = 2700
    Muestra1.ColWidth(5) = 1200
    Muestra1.ColWidth(6) = 1200
    Muestra1.ColWidth(7) = 1200
    
    Muestra1.Height = 4575
    Muestra1.Left = 600
    Muestra1.Top = 120
    Muestra1.Width = 10500
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Fecha"
    
    Muestra1.Col = 2
    Muestra1.Text = "Tipo"
    
    Muestra1.Col = 3
    Muestra1.Text = "Numero"
    
    Muestra1.Col = 4
    Muestra1.Text = "Observaciones"
    
    Muestra1.Col = 5
    Muestra1.Text = "Partida"
    
    Muestra1.Col = 6
    Muestra1.Text = "Cantidad"
    
    Muestra1.Col = 7
    Muestra1.Text = "Saldo"
    
    Muestra1.Visible = True
    
    Renglon = 0
    XStock = 0
    XPedido = 0
    
    If Left$(XProducto, 2) = "DY" Or Left$(XProducto, 2) = "DS" Or Left$(XProducto, 2) = "DQ" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
        
    Select Case WTipopro
        Case "M"
            WArti = Left$(XProducto, 3) + Right$(XProducto, 7)
            
            XParam = "'" + WArti + "','" _
                 + WArti + "'"
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
                    
                            If rstLaudo!Articulo = WArti Then
                            
                                XSaldo = rstLaudo!Saldo
                                Call Redondeo(XSaldo)
                                
                                If XSaldo <> 0 Then
                            
                                    Renglon = Renglon + 1
                                    Muestra1.Row = Renglon
                            
                                    Muestra1.Col = 1
                                    Muestra1.Text = rstLaudo!Fecha
                        
                                    Muestra1.Col = 2
                                    Muestra1.Text = "Laudo"
                                    
                                    Muestra1.Col = 3
                                    Muestra1.Text = rstLaudo!Laudo
                                    
                                    Muestra1.Col = 5
                                    Muestra1.Text = rstLaudo!Laudo
                            
                                    Muestra1.Col = 6
                                    Muestra1.Text = Pusing("###,###.##", Str$(rstLaudo!Liberada))
                                    
                                    Muestra1.Col = 7
                                    Muestra1.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
                rstLaudo.Close
            End If
            
            XParam = "'" + WArti + "','" _
                        + WArti + "'"
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
                
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WMovi = rstMovguia!Movi
                                    
                                If rstMovguia!Tipo = "M" And WMovi = "E" And WSaldo <> 0 Then
                                
                                    Renglon = Renglon + 1
                                    Muestra1.Row = Renglon
                                
                                    Muestra1.Col = 1
                                    Muestra1.Text = rstMovguia!Fecha
                                
                                    Muestra1.Col = 2
                                    Muestra1.Text = "Guia"
                
                                    Muestra1.Col = 3
                                    Muestra1.Text = rstMovguia!Codigo
                        
                                    Muestra1.Col = 4
                                    WTipomov = rstMovguia!Tipomov
                                    Select Case WTipomov
                                        Case 1
                                            Muestra1.Text = "Recep. Surfactan"
                                        Case 2
                                            Muestra1.Text = "Recep. Pellital"
                                        Case 3
                                            Muestra1.Text = "Recep. Surfactan II"
                                        Case 4
                                            Muestra1.Text = "Recep. Pellital II"
                                        Case 5
                                            Muestra1.Text = "Recep. Surfactan III"
                                        Case 6
                                            Muestra1.Text = "Recep. Surfactan IV"
                                        Case 7
                                            Muestra1.Text = "Recep. Surfactan V"
                                        Case 8
                                            Muestra1.Text = "Recep. Pellital V"
                                        Case 9
                                            Muestra1.Text = "Recep. Pellital IV"
                                        Case 10
                                            Muestra1.Text = "Recep. Surfactan VI"
                                        Case 11
                                            Muestra1.Text = "Recep. Surfactan VII"
                                        Case Else
                                    End Select
                        
                                    Muestra1.Col = 5
                                    Muestra1.Text = rstMovguia!Lote
                            
                                    Muestra1.Col = 6
                                    WCantidad = IIf(IsNull(rstMovguia!Cantidad), "0", rstMovguia!Cantidad)
                                    Muestra1.Text = Pusing("###,###.##", Str$(WCantidad))
                                
                                    Muestra1.Col = 7
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Muestra1.Text = Pusing("###,###.##", Str$(WSaldo))
                        
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
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                        
                            If XSaldo <> 0 Then
                        
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                            
                                Muestra1.Col = 1
                                Muestra1.Text = rstHoja!Fecha
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                            
                                Muestra1.Col = 2
                                Muestra1.Text = "Hoja"
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                            
                                Muestra1.Col = 3
                                Muestra1.Text = rstHoja!Hoja
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                            
                                Muestra1.Col = 4
                                Muestra1.Text = ""
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                        
                                Muestra1.Col = 5
                                Muestra1.Text = rstHoja!Hoja
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                            
                                Muestra1.Col = 6
                                Muestra1.Text = Pusing("###,###.##", Str$(rstHoja!Real))
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
                                End If
                                
                                Muestra1.Col = 7
                                Muestra1.Text = Pusing("###,###.##", Str$(rstHoja!Saldo))
                                If WMarcaVencida = "S" Then
                                    Muestra1.CellBackColor = &HC0FFFF
                                End If
                                If WMarcaVencida = "V" Then
                                    Muestra1.CellBackColor = &HFF&
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
                
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WMovi = rstMovguia!Movi
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                    
                                If rstMovguia!Tipo = "T" And WMovi = "E" And WSaldo <> 0 Then
                                
                                    Renglon = Renglon + 1
                                    Muestra1.Row = Renglon
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                                
                                    Muestra1.Col = 1
                                    Muestra1.Text = rstMovguia!Fecha
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                                
                                    Muestra1.Col = 2
                                    Muestra1.Text = "Guia"
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                
                                    Muestra1.Col = 3
                                    Muestra1.Text = rstMovguia!Codigo
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                                    
                                    Muestra1.Col = 4
                                    WTipomov = rstMovguia!Tipomov
                                    Select Case WTipomov
                                        Case 1
                                            Muestra1.Text = "Recep. Surfactan"
                                        Case 2
                                            Muestra1.Text = "Recep. Pellital"
                                        Case 3
                                            Muestra1.Text = "Recep. Surfactan II"
                                        Case 4
                                            Muestra1.Text = "Recep. Pellital II"
                                        Case 5
                                            Muestra1.Text = "Recep. Surfactan III"
                                        Case 6
                                            Muestra1.Text = "Recep. Surfactan IV"
                                        Case 7
                                            Muestra1.Text = "Recep. Surfactan V"
                                        Case 8
                                            Muestra1.Text = "Recep. Pellital V"
                                        Case 9
                                            Muestra1.Text = "Recep. Pellital IV"
                                        Case 10
                                            Muestra1.Text = "Recep. Surfactan VI"
                                        Case 11
                                            Muestra1.Text = "Recep. Surfactan VII"
                                        Case Else
                                    End Select
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                        
                                    Muestra1.Col = 5
                                    Muestra1.Text = rstMovguia!Lote
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                            
                                    Muestra1.Col = 6
                                    WCantidad = IIf(IsNull(rstMovguia!Cantidad), "0", rstMovguia!Cantidad)
                                    Muestra1.Text = Pusing("###,###.##", Str$(WCantidad))
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
                                    End If
                                
                                    Muestra1.Col = 7
                                    WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                    Muestra1.Text = Pusing("###,###.##", Str$(WSaldo))
                                    If WMarcaVencida = "S" Then
                                        Muestra1.CellBackColor = &HC0FFFF
                                    End If
                                    If WMarcaVencida = "V" Then
                                        Muestra1.CellBackColor = &HFF&
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
    
    Call Conecta_Empresa
    
    Muestra1.Row = 1
    Muestra1.Col = 1
    Muestra1.TopRow = 1

End Sub

Private Sub Muestra1_Click()
    If AuxiliarIngresoII.Visible = True Then
        AuxiIngresoIV.Text = Muestra1.TextMatrix(Muestra1.Row, 5)
        AuxiIngresoIV.SetFocus
    End If
    Muestra1.Visible = False
End Sub

Private Sub Muestra1_dblClick()
    If AuxiliarIngresoII.Visible = True Then
        AuxiIngresoIV.Text = Muestra1.TextMatrix(Muestra1.Row, 5)
        AuxiIngresoIV.SetFocus
    End If
    Muestra1.Visible = False
End Sub

Private Sub StkPedido_Click()

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

    Muestra2.Height = 4575
    Muestra2.Left = 800
    Muestra2.Top = 120
    Muestra2.Width = 10700
    
    Muestra2.Clear
    Muestra2.Row = 0
    
    Muestra2.Col = 1
    Muestra2.Text = "Pedido"
    
    Muestra2.Col = 2
    Muestra2.Text = "Cliente"
    
    Muestra2.Col = 3
    Muestra2.Text = "Razon"
    
    Muestra2.Col = 4
    Muestra2.Text = "Fecha"
    
    Muestra2.Col = 5
    Muestra2.Text = "Pedida"
    
    Muestra2.Col = 6
    Muestra2.Text = "Entregada"
    
    Muestra2.Col = 7
    Muestra2.Text = "Saldo"
    
    Muestra2.ColWidth(0) = 100
    Muestra2.ColWidth(1) = 1200
    Muestra2.ColWidth(2) = 1200
    Muestra2.ColWidth(3) = 2500
    Muestra2.ColWidth(4) = 1400
    Muestra2.ColWidth(5) = 1300
    Muestra2.ColWidth(6) = 1300
    Muestra2.ColWidth(7) = 1300
    
    Muestra2.Visible = True

    If Left$(XProducto, 2) = "DY" Or Left$(XProducto, 2) = "DS" Or Left$(XProducto, 2) = "DQ" Then
        WTipopro = "M"
            Else
        WTipopro = "T"
    End If
        
    Select Case WTipopro
        Case "M"
            WArti = Left$(XProducto, 3) + Right$(XProducto, 7)
            Renglon = 0
            spPedido = "ListaPedidoTerminado " + "'" + Termi.Text + "'"
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
                        If XPed <> 0 Then
                        If Pedido.Text <> rstPedido!Pedido Then
                            Renglon = Renglon + 1
                            Muestra2.Row = Renglon
                    
                            Muestra2.Col = 1
                            Muestra2.Text = rstPedido!Pedido
                    
                            Muestra2.Col = 2
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 4
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 5
                            Muestra2.Text = Pusing("###,###.##", Str$(rstPedido!Cantidad))
                            
                            Muestra2.Col = 6
                            Muestra2.Text = Pusing("###,###.##", Str$(rstPedido!Facturado))
                            
                            Muestra2.Col = 7
                            Muestra2.Text = Pusing("###,###.##", Str$(XPed))
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
            Renglon = 0
            spPedido = "ListaPedidoTerminado " + "'" + Termi.Text + "'"
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
                        If XPed <> 0 Then
                        If Pedido.Text <> rstPedido!Pedido Then
                            Renglon = Renglon + 1
                            Muestra2.Row = Renglon
                    
                            Muestra2.Col = 1
                            Muestra2.Text = rstPedido!Pedido
                    
                            Muestra2.Col = 2
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 4
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 5
                            Muestra2.Text = Pusing("###,###.##", Str$(rstPedido!Cantidad))
                            
                            Muestra2.Col = 6
                            Muestra2.Text = Pusing("###,###.##", Str$(rstPedido!Facturado))
                            
                            Muestra2.Col = 7
                            Muestra2.Text = Pusing("###,###.##", Str$(XPed))
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
    
    For dada = 1 To Renglon
    
        Muestra2.Row = dada
                        
        Muestra2.Col = 2
        WCliente = Muestra2.Text
    
        spCliente = "ConsultaClienteRazon " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra2.Col = 3
            Muestra2.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next dada
    
    Call Conecta_Empresa
    
    Muestra2.Row = 1
    Muestra2.Col = 1
    Muestra2.TopRow = 1
    
End Sub

Private Sub Muestra2_Click()
    Muestra2.Visible = False
End Sub

Private Sub Muestra2_dblClick()
    Muestra2.Visible = False
End Sub

Private Sub Muestra3_Click()
    Muestra3.Visible = False
End Sub

Private Sub Muestra4_Click()
    Muestra4.Visible = False
    AuxiIngresoVII.SetFocus
End Sub

Private Sub StkProduccion_Click()

    Muestra3.Height = 4575
    Muestra3.Left = 200
    Muestra3.Top = 120
    Muestra3.Width = 12700
    
    Muestra3.Clear
    Muestra3.Row = 0
    
    Muestra3.Col = 1
    Muestra3.Text = "Solicitud"
    
    Muestra3.Col = 2
    Muestra3.Text = "Observaciones"
    
    Muestra3.Col = 3
    Muestra3.Text = "Articulo"
    
    Muestra3.Col = 4
    Muestra3.Text = "Fecha"
    
    Muestra3.Col = 5
    Muestra3.Text = "Pedida"
    
    Muestra3.Col = 6
    Muestra3.Text = "Entregada"
    
    Muestra3.Col = 7
    Muestra3.Text = "Saldo"
    
    Muestra3.ColWidth(0) = 100
    Muestra3.ColWidth(1) = 1000
    Muestra3.ColWidth(2) = 3500
    Muestra3.ColWidth(3) = 1500
    Muestra3.ColWidth(4) = 1400
    Muestra3.ColWidth(5) = 1300
    Muestra3.ColWidth(6) = 1300
    Muestra3.ColWidth(7) = 1300
    
    Muestra3.Visible = True
    
    
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSolicitud"
    ZSql = ZSql + " Where CargaSolicitud.Articulo = " + "'" + Termi.Text + "'"
    ZSql = ZSql + " and CargaSolicitud.Saldo > 0"
    ZSql = ZSql + " Order by CargaSolicitud.Clave"
    spCargaSolicitud = ZSql
    Set rstCargaSolicitud = db.OpenRecordset(spCargaSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSolicitud.RecordCount > 0 Then
        With rstCargaSolicitud
            .MoveFirst
            If .NoMatch = False Then
    
            Do
    
                If .EOF = True Then
                    Exit Do
                End If
        
                Renglon = Renglon + 1
                Muestra3.Row = Renglon
        
                Muestra3.Col = 1
                Muestra3.Text = rstCargaSolicitud!Solicitud
        
                Muestra3.Col = 2
                Muestra3.Text = rstCargaSolicitud!Observaciones
                
                Muestra3.Col = 3
                Muestra3.Text = rstCargaSolicitud!Articulo
    
                Muestra3.Col = 4
                Muestra3.Text = rstCargaSolicitud!Fecha
                
                Muestra3.Col = 5
                Muestra3.Text = Pusing("###,###.##", Str$(rstCargaSolicitud!Cantidad))
                
                Muestra3.Col = 6
                Muestra3.Text = Pusing("###,###.##", Str$(rstCargaSolicitud!Entregado))
                
                Muestra3.Col = 7
                Muestra3.Text = Pusing("###,###.##", Str$(rstCargaSolicitud!Saldo))
        
                .MoveNext
        
                If .EOF = True Then
                    Exit Do
                End If
        
            Loop
            End If
        End With
    End If
    
    For dada = 1 To Renglon
    
        Muestra3.Row = dada
                        
        Muestra3.Col = 2
        WCliente = Muestra3.Text
    
        spCliente = "ConsultaClienteRazon " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Muestra3.Col = 3
            Muestra3.Text = rstCliente!Razon
            rstCliente.Close
        End If
        
    Next dada
    
    Call Conecta_Empresa
    
    Muestra3.Row = 1
    Muestra3.Col = 1
    Muestra3.TopRow = 1

End Sub

Private Sub VerMp_Click()

    Muestra4.Height = 4575
    Muestra4.Left = 200
    Muestra4.Top = 120
    Muestra4.Width = 12700
    
    Muestra4.Clear
    Muestra4.Row = 0
    
    Muestra4.Col = 1
    Muestra4.Text = "Tipo"
    
    Muestra4.Col = 2
    Muestra4.Text = "P.Terminado"
    
    Muestra4.Col = 3
    Muestra4.Text = "M.Prima"
    
    Muestra4.Col = 4
    Muestra4.Text = "Descripcion"
    
    Muestra4.Col = 5
    Muestra4.Text = "Cantidad"
    
    Muestra4.Col = 6
    Muestra4.Text = "Stock II"
    
    Muestra4.Col = 7
    Muestra4.Text = "Stock V"
    
    Muestra4.Col = 8
    Muestra4.Text = "Diferecia"
    
    Muestra4.ColWidth(0) = 100
    Muestra4.ColWidth(1) = 1000
    Muestra4.ColWidth(2) = 1400
    Muestra4.ColWidth(3) = 1400
    Muestra4.ColWidth(4) = 2800
    Muestra4.ColWidth(5) = 1100
    Muestra4.ColWidth(6) = 1100
    Muestra4.ColWidth(7) = 1100
    Muestra4.ColWidth(8) = 1100
    
    Muestra4.ColAlignment(5) = flexAlignRightCenter
    Muestra4.ColAlignment(6) = flexAlignRightCenter
    Muestra4.ColAlignment(7) = flexAlignRightCenter
    Muestra4.ColAlignment(8) = flexAlignRightCenter
    
    Muestra4.Visible = True
    
    Renglon = 0
    ZZCantidad = 0
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Termi.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
        
                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
        
                        Renglon = Renglon + 1
                        
                        Muestra4.TextMatrix(Renglon, 1) = rstComposicion!Tipo
                    
                        Muestra4.TextMatrix(Renglon, 2) = rstComposicion!Articulo2
                        Auxi2 = rstComposicion!Articulo2
                    
                        If rstComposicion!Articulo1 = "  -   -  " Then
                            Muestra4.TextMatrix(Renglon, 3) = "  -   -   "
                            Auxi1 = "  -   -   "
                                Else
                            Muestra4.TextMatrix(Renglon, 3) = rstComposicion!Articulo1
                            Auxi1 = rstComposicion!Articulo1
                        End If
                    
                        Cantidad = Str$(rstComposicion!Cantidad * Val(AuxiIngresoVII))
                        ZZCantidad = rstComposicion!Cantidad
                        
                        Auxi2 = Cantidad
                        Auxi2 = Pusing("###,###.##", Auxi2)
                        Muestra4.TextMatrix(Renglon, 5) = Auxi2
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    For ZZCiclo = 1 To Renglon
    
        ZZTipo = Muestra4.TextMatrix(ZZCiclo, 1)
        ZZTerminado = Muestra4.TextMatrix(ZZCiclo, 2)
        ZZArticulo = Muestra4.TextMatrix(ZZCiclo, 3)
    
        Select Case ZZTipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + ZZTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Muestra4.TextMatrix(ZZCiclo, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Muestra4.TextMatrix(ZZCiclo, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                XEmpresa = WEmpresa
        
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZStockII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    ZStockII = 0
                End If
                Auxi2 = Str$(ZStockII)
                Auxi2 = Pusing("###,###.##", Auxi2)
                Muestra4.TextMatrix(ZZCiclo, 6) = Auxi2
                
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                      
                spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZStockIII = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    ZStockIII = 0
                End If
                Auxi2 = Str$(ZStockIII)
                Auxi2 = Pusing("###,###.##", Auxi2)
                Muestra4.TextMatrix(ZZCiclo, 7) = Auxi2
    
                Auxi2 = ZStockII + ZStockIII - Val(Muestra4.TextMatrix(ZZCiclo, 5))
                Auxi2 = Pusing("###,###.##", Auxi2)
                Muestra4.TextMatrix(ZZCiclo, 8) = Auxi2
    
                Call Conecta_Empresa
                
            Case Else
        End Select
    
    Next ZZCiclo
    
    Muestra4.Row = 1
    Muestra4.Col = 1
    Muestra4.TopRow = 1
    
End Sub

Private Sub WStock1_Click()
    WProceso = 1
    Call Muestra_Ficha
End Sub

Private Sub WStock1_dblClick()
    WProceso = 1
    Call Muestra_Ficha
End Sub

Private Sub WStock2_Click()
    WProceso = 2
    Call Muestra_Ficha
End Sub

Private Sub WStock2_dblClick()
    WProceso = 2
    Call Muestra_Ficha
End Sub

Private Sub WStock3_Click()
    WProceso = 3
    Call Muestra_Ficha
End Sub

Private Sub WStock3_dblClick()
    WProceso = 3
    Call Muestra_Ficha
End Sub

Private Sub WStock4_Click()
    WProceso = 4
    Call Muestra_Ficha
End Sub


Private Sub AyudaOrden()

    Orden.Height = 3000
    Orden.Left = 5980
    Orden.Top = 1200
    Orden.Width = 3500
    
    Orden.Clear
    
    Orden.Row = 0
    Orden.Col = 1
    Orden.Text = "Procedimiento"
    
    Orden.ColWidth(0) = 100
    Orden.ColWidth(1) = 3000
    
    Orden.Visible = True
    
    Orden.Row = 1
    Orden.Col = 1
    Orden.Text = "Stock"
    
    Orden.Row = 2
    Orden.Col = 1
    Orden.Text = "Produccion"
    
    Orden.Row = 3
    Orden.Col = 1
    Orden.Text = "Parcial"
    
    Orden.Row = 4
    Orden.Col = 1
    Orden.Text = "Varios"
    
    Orden.Row = 1
    Orden.Col = 1
    Orden.TopRow = 1
    Orden.SetFocus
    
End Sub

Private Sub Orden_Click()
    Call Orden_dblclick
End Sub

Private Sub Orden_dblclick()

    Muestra.TopRow = Lugar1
    Muestra.Row = Lugar2
    Muestra.Col = Lugar3
    Muestra.Text = Orden.Text
    Orden.Visible = False
    
    Select Case Orden.Row
        Case 1, 3
            AuxiliarIngresoII.Height = 1335
            AuxiliarIngresoII.Left = 1560
            AuxiliarIngresoII.Top = 1920
            AuxiliarIngresoII.Width = 7695
            AuxiliarIngresoII.Visible = True
            AuxiIngresoIV.Text = Muestra.TextMatrix(Muestra.Row, 6)
            AuxiIngresoV.Text = Muestra.TextMatrix(Muestra.Row, 7)
            AuxiIngresoIV.SetFocus
            
        Case Else
            AuxiliarIngresoIII.Height = 1335
            AuxiliarIngresoIII.Left = 1560
            AuxiliarIngresoIII.Top = 1920
            AuxiliarIngresoIII.Width = 7695
            AuxiliarIngresoIII.Visible = True
            AuxiIngresoVI.Text = Muestra.TextMatrix(Muestra.Row, 7)
            AuxiIngresoVII.Text = Muestra.TextMatrix(Muestra.Row, 6)
            AuxiIngresoVI.SetFocus
            
    End Select
End Sub

Private Sub AuxiIngreso_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AuxiIngreso.Text = Pusing("###,###.##", AuxiIngreso.Text)
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub AuxiIngresoIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem AuxiIngresoIV.Text = Pusing("###,###.##", AuxiIngresoIV.Text)
        AuxiIngresoV.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub AuxiIngresoV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Muestra.TopRow = Lugar1
        Muestra.Row = Lugar2
        Muestra.Col = 6
        Muestra.Text = AuxiIngresoIV.Text
        Muestra.Col = 7
        Muestra.Text = AuxiIngresoV.Text
        AuxiliarIngresoII.Visible = False
    End If
End Sub

Private Sub AuxiIngresoVI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AuxiIngresoVII.SetFocus
    End If
End Sub

Private Sub AuxiIngresoVII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Muestra.TopRow = Lugar1
        Muestra.Row = Lugar2
        Muestra.Col = 7
        Muestra.Text = AuxiIngresoVI.Text
        AuxiliarIngresoIII.Visible = False
    End If
End Sub

Private Sub Graba_Click()

    Rem For Ciclo = 1 To 99
    Rem     If Val(Muestra.TextMatrix(Ciclo, 4)) <> 0 Then
    Rem         If Muestra.TextMatrix(Ciclo, 5) = "" Then
    Rem             Exit Sub
    Rem         End If
    Rem     End If
    Rem Next Ciclo
    
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
    
    Traspa = "N"
    
    For Ciclo = 1 To 99
        If Muestra.TextMatrix(Ciclo, 1) <> "" Then
        
            WProc1 = "2"
            
            If Muestra.TextMatrix(Ciclo, 5) = "Stock" Then
                WProc2 = "1"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "" Then
                WProc2 = "1"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Produccion" Then
                WProc2 = "2"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Parcial" Then
                Traspa = "S"
                WProc2 = "3"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Varios" Then
                WProc2 = "4"
            End If
            
            WClave = ClavePedido(Ciclo)
            WProc3 = Transfe(Ciclo, 1)
            WProc4 = Transfe(Ciclo, 2)
            WProc5 = Transfe(Ciclo, 3)
            WProc6 = ""
            WProc7 = Muestra.TextMatrix(Ciclo, 7)
            
            XParam = "'" + WClave + "','" _
                         + WProc1 + "','" _
                         + WProc2 + "','" _
                         + WProc3 + "','" _
                         + WProc4 + "','" _
                         + WProc5 + "','" _
                         + WProc6 + "','" _
                         + WProc7 + "'"

            spPedido = "ModificaPedidoProceso2 " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " Partida = " + "'" + Muestra.TextMatrix(Ciclo, 6) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
    Next Ciclo
    
    WMarca = "2"
    XParam = "'" + Pedido.Text + "','" _
                 + WMarca + "'"
                                           
    spPedido = "ModificaPedidoProceso1 " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    If Traspa = "S" Then
    
        For Ciclo1 = 1 To 5
            For Ciclo2 = 1 To 5
            
                Pasa = 0
                Xlugar = 0
                WNumeroSolGuia = ""
                
                For ciclo3 = 1 To 99
                    If Val(Transfe(ciclo3, 2)) = Ciclo1 And Val(Transfe(ciclo3, 3)) = Ciclo2 Then
                    
                        If Pasa = 0 Then
                            Pasa = 1
                            spSolGuia = "ListaSolguiaNumero "
                            Set rstSolGuia = db.OpenRecordset(spSolGuia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstSolGuia.RecordCount > 0 Then
                                With rstSolGuia
                                    .MoveLast
                                    Do
                                        WNumeroSolGuia = rstSolGuia!Codigo + 1
                                        Exit Do
                                    Loop
                                End With
                                rstSolGuia.Close
                                    Else
                                WNumeroSolGuia = "1"
                            End If
                        End If
        
                        Xlugar = Xlugar + 1
                        
                        WTipo = "T"
                        WTerminado = Transfe(ciclo3, 5)
                        WArticulo = "  -   -   "
                        WCantidad = Transfe(ciclo3, 1)
                    
                        Auxi = Str$(Xlugar)
                        Call Ceros(Auxi, 2)
                        
                        Auxi1 = WNumeroSolGuia
                        Call Ceros(Auxi1, 6)
                
                        WDesde = Transfe(ciclo3, 2)
                        WHasta = Transfe(ciclo3, 3)
                
                        WCodigo = WNumeroSolGuia
                        WRenglon = Str$(Xlugar)
                        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WObservaciones = "P/" + Pedido.Text + " " + "(" + Left$(DesCliente.Caption, 30) + ")"
                        WClave = Auxi1 + Auxi
                        WMarca = "N"
                        WAviso = "0"
                        WUsuario = ""

                        XParam = "'" + WClave + "','" _
                            + WCodigo + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WTipo + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WCantidad + "','" _
                            + WFechaord + "','" _
                            + WObservaciones + "','" _
                            + WDesde + "','" _
                            + WHasta + "','" _
                            + WMarca + "','" _
                            + WUsuario + "','" _
                            + WAviso + "'"
                         
                        spSolGuia = "AltaSolguia " + XParam
                        Set rstSolGuia = db.OpenRecordset(spSolGuia, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                Next ciclo3
            Next Ciclo2
        Next Ciclo1
        
    End If
    
    Call Conecta_Empresa

    With rstEmpresa
        .Close
    End With
    PrgPedCentroPelli.Hide
    Unload Me
    PrgCentro.Show

End Sub

