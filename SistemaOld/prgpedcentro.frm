VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPedCentro 
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
   Begin VB.Frame AuxiliarIngresoIII 
      Height          =   1335
      Left            =   1560
      TabIndex        =   54
      Top             =   3240
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox AuxiIngresoVI 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   55
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   315
         Left            =   480
         TabIndex        =   56
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame AuxiliarIngresoII 
      Height          =   1335
      Left            =   1560
      TabIndex        =   49
      Top             =   1920
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox AuxiIngresoV 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   53
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox AuxiIngresoIV 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   50
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   315
         Left            =   480
         TabIndex        =   52
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero de Partida"
         Height          =   315
         Left            =   480
         TabIndex        =   51
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame AuxiliarIngresoI 
      Height          =   2895
      Left            =   7200
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox AuxiIngresoVII 
         Height          =   315
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   57
         Top             =   1680
         Width           =   5055
      End
      Begin VB.CommandButton CancelaAuxiliarIngresoI 
         Caption         =   "Cancela"
         Height          =   375
         Left            =   2760
         TabIndex        =   48
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton ConfirmaAuxiliarIngresoI 
         Caption         =   "Confirma"
         Height          =   375
         Left            =   600
         TabIndex        =   47
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ComboBox AuxiIngresoII 
         Height          =   360
         Left            =   2640
         TabIndex        =   44
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox AuxiIngresoI 
         Height          =   360
         Left            =   2640
         TabIndex        =   43
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox AuxiIngreso 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   315
         Left            =   240
         TabIndex        =   58
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta Planta"
         Height          =   360
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde Planta"
         Height          =   360
         Left            =   240
         TabIndex        =   45
         Top             =   750
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad a Transferir"
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   3015
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Orden 
      Height          =   1215
      Left            =   1920
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
      _Version        =   327680
      Rows            =   100
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
      _Version        =   327680
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
      _Version        =   327680
      Rows            =   100
      Cols            =   8
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
      TabIndex        =   36
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
      _Version        =   327680
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
         TabIndex        =   38
         Top             =   240
         Width           =   3015
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
         TabIndex        =   62
         Top             =   2880
         Width           =   1815
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
         Left            =   1920
         TabIndex        =   61
         Top             =   2880
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
         Left            =   120
         TabIndex        =   60
         Top             =   3120
         Width           =   1695
      End
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
         Left            =   1920
         TabIndex        =   59
         Top             =   3120
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
         TabIndex        =   37
         Top             =   600
         Width           =   3015
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
         Left            =   1920
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
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
         TabIndex        =   34
         Top             =   2640
         Width           =   1695
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
         Top             =   2400
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
         Top             =   2400
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
         Top             =   2160
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
         Top             =   2160
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
         Top             =   1920
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
         Top             =   1680
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
         Top             =   1920
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
         Top             =   1680
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
         Top             =   1440
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
         Top             =   1440
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
Attribute VB_Name = "PrgPedCentro"
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
Dim XParam As String
Dim ClavePedido(100)
Dim WProceso As Integer
Dim WSaldo As Double
Dim WNumeroSolGuia As String
Dim XPedido As Double

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgPedCentro.Hide
    Unload Me
    PrgCentro.Show
End Sub

Private Sub CancelaAuxiliarIngresoI_Click()
    Muestra.TopRow = Lugar1
    Muestra.Row = Lugar2
    Muestra.Col = 5
    Muestra.Text = ""
    Muestra.Col = 6
    Muestra.Text = ""
    Muestra.Col = 7
    Muestra.Text = ""
    Transfe(Muestra.Row, 1) = ""
    Transfe(Muestra.Row, 2) = ""
    Transfe(Muestra.Row, 3) = ""
    Transfe(Muestra.Row, 4) = ""
    Transfe(Muestra.Row, 5) = ""
    AuxiliarIngresoI.Visible = False
End Sub

Private Sub ConfirmaAuxiliarIngresoI_Click()
    If Val(AuxiIngreso.Text) <> 0 And AuxiIngresoI.ListIndex <> 0 And AuxiIngresoII.ListIndex <> 0 Then
        Muestra.TopRow = Lugar1
        Muestra.Row = Lugar2
        Muestra.Col = 6
        Muestra.Text = AuxiIngreso.Text
        Muestra.Col = 7
        Muestra.Text = AuxiIngresoVII.Text
        Transfe(Muestra.Row, 1) = AuxiIngreso.Text
        Transfe(Muestra.Row, 2) = Str$(AuxiIngresoI.ListIndex)
        Transfe(Muestra.Row, 3) = Str$(AuxiIngresoII.ListIndex)
        Transfe(Muestra.Row, 5) = Muestra.TextMatrix(Muestra.Row, 1)
        AuxiliarIngresoI.Visible = False
    End If
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
        
        If Left$(XProducto, 2) = "PT" Or Left$(XProducto, 2) = "YQ" Or Left$(XProducto, 2) = "YF" Or Left$(XProducto, 2) = "YP" Then
            WTipopro = "T"
                Else
            WTipopro = "M"
        End If
    
        Call Stock_Consolidado
        
        XStock = (Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(WStock6.Caption) + Val(WStock7.Caption))
        Stock.Caption = Pusing("###,###.##", Str$(XStock))
        aa = Muestra.TextMatrix(Muestra.Row, 4)
        StkPedido.Caption = Pusing("###,###.##", Str$(XPedido - Val(Muestra.TextMatrix(Muestra.Row, 4))))
        Disponible.Caption = Pusing("###,###.##", Str$(XStock - Val(StkPedido.Caption)))
    
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
    
    XEmpresa = WEmpresa
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        
        Stock1.Caption = "Pta I"
        Stock2.Caption = "Pta II"
        Stock3.Caption = "Pta III"
        Stock4.Caption = "Pta IV"
        Stock5.Caption = "Pta V"
        Stock6.Caption = "Pta VI"
        Stock7.Caption = "Pta VII"
        
        If Left$(Termi.Text, 2) = "PT" Or Left$(Termi.Text, 2) = "YQ" Or Left$(Termi.Text, 2) = "YF" Or Left$(Termi.Text, 2) = "YP" Then
            WTipopro = "T"
                Else
            WTipopro = "M"
        End If
        WArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock1.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    XPedido = IIf(IsNull(rstArticulo!Venta), "0", rstArticulo!Venta)
                    rstArticulo.Close
                        Else
                    WStock1.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    XPedido = IIf(IsNull(rstTerminado!Pedido), "0", rstTerminado!Pedido)
                    rstTerminado.Close
                        Else
                    WStock1.Caption = "0"
                End If
        End Select
            
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock2.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock2.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock2.Caption = "0"
                End If
        End Select
                    
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
              
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock3.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock3.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock3.Caption = "0"
                End If
        End Select
                
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock4.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock4.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock4.Caption = "0"
                End If
        End Select
                
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock5.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock5.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock5.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock5.Caption = "0"
                End If
        End Select
                
        WEmpresa = "0010"
        txtOdbc = "Empresa10"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock6.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock6.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock6.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock6.Caption = "0"
                End If
        End Select
                
        WEmpresa = "0011"
        txtOdbc = "Empresa11"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        Select Case WTipopro
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WStock7.Caption = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                        Else
                    WStock7.Caption = "0"
                End If
                        
            Case Else
                spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WStock7.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                        Else
                    WStock7.Caption = "0"
                End If
        End Select
        
            Else
            
        
        Stock1.Caption = "Pta I"
        Stock2.Caption = "Pta II"
        Stock3.Caption = "Pta V"
        Stock4.Caption = "Pta IV"
        Stock5.Caption = ""
        Stock6.Caption = ""
        Stock7.Caption = ""
        
        If Left$(Termi.Text, 2) = "PT" Or Left$(Termi.Text, 2) = "YQ" Or Left$(Termi.Text, 2) = "YF" Or Left$(Termi.Text, 2) = "YP" Then
            WTipopro = "T"
                Else
            WTipopro = "M"
        End If
        WArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
        
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
        spTerminado = "Consultaterminado " + "'" + Termi.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
            XPedido = IIf(IsNull(rstTerminado!Pedido), "0", rstTerminado!Pedido)
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
        
        WStock5.Caption = "0"
        WStock6.Caption = "0"
        WStock7.Caption = "0"
        
            
    End If
    
    Call Conecta_Empresa
            
    WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
    WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
    WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
    WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
    WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
    WStock6.Caption = Pusing("###,###.##", WStock6.Caption)
    WStock7.Caption = Pusing("###,###.##", WStock7.Caption)

End Sub

Private Sub Muestra_Ficha()

    XEmpresa = WEmpresa
    XProducto = Termi.Text
    
    Select Case WProceso
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
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
    Muestra1.Visible = False
End Sub

Private Sub Muestra1_dblClick()
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

Private Sub WStock5_Click()
    WProceso = 5
    Call Muestra_Ficha
End Sub

Private Sub WStock6_Click()
    WProceso = 6
    Call Muestra_Ficha
End Sub

Private Sub WStock7_Click()
    WProceso = 7
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
    Orden.Text = "Produccion Pta."
    
    Orden.Row = 3
    Orden.Col = 1
    Orden.Text = "Transferencia"
    
    Orden.Row = 4
    Orden.Col = 1
    Orden.Text = "Produccion Pta. (Fecha)"
    
    Orden.Row = 5
    Orden.Col = 1
    Orden.Text = "Parcial (Produccion Fecha)"
    
    Orden.Row = 6
    Orden.Col = 1
    Orden.Text = "Produccion P1 (MP)"
    
    Orden.Row = 7
    Orden.Col = 1
    Orden.Text = "FMP"
    
    Orden.Row = 8
    Orden.Col = 1
    Orden.Text = "Kgrs (Pellital)"
    
    Orden.Row = 9
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
        Case 3
            AuxiIngresoI.Clear
            AuxiIngresoI.AddItem ""
            AuxiIngresoI.AddItem "Surfactan"
            AuxiIngresoI.AddItem "Surfactan II"
            AuxiIngresoI.AddItem "Surfactan III"
            AuxiIngresoI.AddItem "Surfactan IV"
            AuxiIngresoI.AddItem "Surfactan V"
            AuxiIngresoI.AddItem "Surfactan VI"
            AuxiIngresoI.AddItem "Surfactan VII"
            
            AuxiIngresoII.Clear
            AuxiIngresoII.AddItem ""
            AuxiIngresoII.AddItem "Surfactan"
            AuxiIngresoII.AddItem "Surfactan II"
            AuxiIngresoII.AddItem "Surfactan III"
            AuxiIngresoII.AddItem "Surfactan IV"
            AuxiIngresoII.AddItem "Surfactan V"
            AuxiIngresoII.AddItem "Surfactan VI"
            AuxiIngresoII.AddItem "Surfactan VII"
            
            AuxiliarIngresoI.Height = 3135
            AuxiliarIngresoI.Left = 2040
            AuxiliarIngresoI.Top = 1920
            AuxiliarIngresoI.Width = 7295
            AuxiliarIngresoI.Visible = True
            
            If Val(Transfe(Muestra.Row, 1)) = 0 Then
                AuxiIngreso.Text = Muestra.TextMatrix(Muestra.Row, 4)
                AuxiIngresoI.ListIndex = 0
                AuxiIngresoII.ListIndex = 1
                    Else
                AuxiIngreso.Text = Transfe(Muestra.Row, 1)
                AuxiIngresoI.ListIndex = Val(Transfe(Muestra.Row, 2))
                AuxiIngresoII.ListIndex = Val(Transfe(Muestra.Row, 3))
            End If
            
            AuxiIngresoVII.Text = Muestra.TextMatrix(Muestra.Row, 7)
            
            AuxiliarIngresoI.Visible = True
            AuxiIngreso.SetFocus
            
        Case 2
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
        AuxiIngresoIV.Text = Pusing("###,###.##", AuxiIngresoIV.Text)
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
            
            If Muestra.TextMatrix(Ciclo, 5) = "Produccion Pta." Then
                WProc2 = "2"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Transferencia" Then
                Traspa = "S"
                WProc2 = "3"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Produccion Pta. (Fecha)" Then
                WProc2 = "4"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Parcial (Produccion Fecha)" Then
                WProc2 = "5"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Produccion P1 (MP)" Then
                WProc2 = "6"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "FMP" Then
                WProc2 = "7"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Kgrs (Pellital)" Then
                WProc2 = "8"
            End If
            
            If Muestra.TextMatrix(Ciclo, 5) = "Varios" Then
                WProc2 = "9"
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
    PrgPedCentro.Hide
    Unload Me
    PrgCentro.Show

End Sub

