VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMuestraPedido 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Pedidos"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Begin VB.Frame Pass 
      Height          =   1575
      Left            =   4080
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   840
         TabIndex        =   48
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   47
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox LoteImpre 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   45
      Text            =   " "
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tanatex"
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   1800
      TabIndex        =   42
      Top             =   3480
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   41
      Top             =   4080
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      TabIndex        =   40
      Top             =   3480
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   2640
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
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2520
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
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2400
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2280
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2400
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4455
      Left            =   120
      TabIndex        =   33
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7858
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.ComboBox MarcaFactura 
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
      Left            =   1320
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
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
      Height          =   495
      Left            =   8640
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton ReImpre 
      Caption         =   "ReImpresion"
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
      Left            =   7200
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Canti5 
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
      Left            =   2160
      TabIndex        =   22
      Text            =   " "
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Canti4 
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
      Left            =   2160
      TabIndex        =   21
      Text            =   " "
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Canti3 
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
      Left            =   2160
      TabIndex        =   20
      Text            =   " "
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Canti2 
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
      Left            =   2160
      TabIndex        =   19
      Text            =   " "
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Canti1 
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
      Left            =   2160
      TabIndex        =   18
      Text            =   " "
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Envase5 
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
      TabIndex        =   17
      Text            =   " "
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Envase4 
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
      TabIndex        =   16
      Text            =   " "
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Envase3 
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
      TabIndex        =   15
      Text            =   " "
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Envase2 
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
      TabIndex        =   14
      Text            =   " "
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Envase1 
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
      TabIndex        =   13
      Text            =   " "
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
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
      Left            =   10080
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   6120
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Cliente 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   8
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
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
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
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
      Height          =   2220
      ItemData        =   "MuestraPedido.frx":0000
      Left            =   3360
      List            =   "MuestraPedido.frx":0007
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   8055
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3000
      TabIndex        =   43
      Top             =   3480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENVASES A ENTREGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Descri5 
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
      Left            =   1200
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Descri4 
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
      Left            =   1200
      TabIndex        =   26
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Descri3 
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
      Left            =   1200
      TabIndex        =   25
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Descri2 
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
      Left            =   1200
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Descri1 
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
      Left            =   1200
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Pedido"
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
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label3 
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
      TabIndex        =   7
      Top             =   480
      Width           =   1095
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
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgMuestraPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WAceptada As String
Private WDirentrega As String
Private WFecEntrega As String
Private WDespago As String
Private WObservaciones As String
Private WVersion As String
Private WTipoPedido As String
Private WMarcaFactura As Integer

Private WImpre(10) As String
Private WEnvase(10) As String
Private Envase(5, 2) As String
Private Auxiliar(100, 14) As String
Private ClavePedido(100) As String
Private bajaLote(5, 2) As String
Private XLote(100, 100) As String
Private XEnvase(100, 6) As String
Private ImpreEnvase(10) As String
Private EmiteCerti(1000, 3) As String
Private CargaEmpresa(10, 2) As String
Private LugarCerti As Integer
Private TipoEnvase(100, 2) As String

Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPago As Recordset
Dim spPago As String

Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstCertificado As Recordset
Dim spCertificado As String

Dim ZAprueba As String
Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 2) As String
Dim ZDescri(10) As String
Dim ZClave1 As String
Dim ZClave2 As String

Dim XParam As String

Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WSaldo4 As Double
Dim WSaldo5 As Double
Dim WSaldo6 As Double
Dim WSaldo7 As Double
Dim WSaldo8 As Double
Dim WSaldo9 As Double
Dim WSaldo10 As Double
Dim WSaldo11 As Double
Dim WSaldo12 As Double

Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim XSaldo4 As String
Dim XSaldo5 As String
Dim XSaldo6 As String
Dim XSaldo7 As String
Dim XSaldo8 As String
Dim XSaldo9 As String
Dim XSaldo10 As String
Dim XSaldo11 As String
Dim XSaldo12 As String

Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer

Dim XCantidad1 As String
Dim xCantidad2 As String

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
Dim XLote6 As String
Dim XCantiLote6 As String
Dim XLote7 As String
Dim XCantiLote7 As String
Dim XLote8 As String
Dim XCantiLote8 As String
Dim XLote9 As String
Dim XCantiLote9 As String
Dim XLote10 As String
Dim XCantiLote10 As String
Dim XLote11 As String
Dim XCantiLote11 As String
Dim XLote12 As String
Dim XCantiLote12 As String

Dim XEnv1 As String
Dim XCantiEnv1 As String
Dim XBultos1 As String
Dim XEnv2 As String
Dim XCantiEnv2 As String
Dim XBultos2 As String
Dim XEnv3 As String
Dim XCantiEnv3 As String
Dim XBultos3 As String
Dim XEnv4 As String
Dim XCantiEnv4 As String
Dim XBultos4 As String
Dim XEnv5 As String
Dim XCantiEnv5 As String
Dim XBultos5 As String
Dim XEnv6 As String
Dim XCantiEnv6 As String
Dim XBultos6 As String
Dim XEnv7 As String
Dim XCantiEnv7 As String
Dim XBultos7 As String
Dim XEnv8 As String
Dim XCantiEnv8 As String
Dim XBultos8 As String
Dim XEnv9 As String
Dim XCantiEnv9 As String
Dim XBultos9 As String
Dim XEnv10 As String
Dim XCantiEnv10 As String
Dim XBultos10 As String
Dim XEnv11 As String
Dim XCantiEnv11 As String
Dim XBultos11 As String
Dim XEnv12 As String
Dim XCantiEnv12 As String
Dim XBultos12 As String

Dim XMes As String
Dim XAno As String

Dim ControlLote(12, 2) As String

Dim WSaldo As Double
Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer

Dim ZZGrilla(100, 15) As String
Dim ZZHoja(100) As String
Dim ZZNumeroHoja As String
Dim ZZZArticulo As String
Dim ZZZPartida As String

Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim XEspecificaciones(100) As String
Dim ZVector(100, 11) As String
Dim WEspecif(100) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub cmdClose_Click()
    PrgMuestraPedido.Hide
    Unload Me
    PrgActualizaPedidoFarma.Show
End Sub

Private Sub Form_Activate()

    If ZZProcesoFactura = 99 And Val(Pedido.Text) <> 0 Then
        Call Pedido_KeyPress(13)
    End If

End Sub



Private Sub Verifica_Certificado()

    WRenglon = 0
    ZAprueba = "S"

    For A = 1 To 99
        
        WRenglon = WRenglon + 1
            
        Articulo = WVector1.TextMatrix(WRenglon, 1)
        
        If Trim(Articulo) <> "" Then
                
            WLugar = WRenglon
                
            WLote1 = XLote(WLugar, 1)
            WLote2 = XLote(WLugar, 3)
            WLote3 = XLote(WLugar, 5)
            WLote4 = XLote(WLugar, 7)
            WLote5 = XLote(WLugar, 9)
            WLote6 = XLote(WLugar, 11)
            WLote7 = XLote(WLugar, 13)
            WLote8 = XLote(WLugar, 15)
            WLote9 = XLote(WLugar, 17)
            WLote10 = XLote(WLugar, 19)
            WLote11 = XLote(WLugar, 21)
            WLote12 = XLote(WLugar, 23)
        
            Rem
            Rem certificado de analisis
            Rem
        
            For ZZCiclo = 1 To 12
                
                Select Case ZZCiclo
                    Case 1
                        WWLote = WLote1
                    Case 2
                        WWLote = WLote2
                    Case 3
                        WWLote = WLote3
                    Case 4
                        WWLote = WLote4
                    Case 5
                        WWLote = WLote5
                    Case 6
                        WWLote = WLote6
                    Case 7
                        WWLote = WLote7
                    Case 8
                        WWLote = WLote8
                    Case 9
                        WWLote = WLote9
                    Case 10
                        WWLote = WLote10
                    Case 11
                        WWLote = WLote11
                    Case Else
                        WWLote = WLote12
                End Select
                
                If Trim(WWLote) <> "" Then
            
                    ZZEntra = "N"
            
                    If Left$(UCase(Articulo), 2) = "PT" Then
                    
                        XCodigo = Val(Mid$(Articulo, 4, 5))
                        If XCodigo >= 0 And XCodigo <= 999 Then
                            XTipoPro = "CO"
                                Else
                            If XCodigo >= 11000 And XCodigo <= 12999 Then
                                XTipoPro = "CO"
                                    Else
                                If XCodigo >= 25000 And XCodigo <= 25999 Then
                                    XTipoPro = "FA"
                                        Else
                                    If XCodigo >= 2300 And XCodigo <= 2399 Then
                                        XTipoPro = "BI"
                                            Else
                                        If XCodigo >= 40000 And XCodigo <= 41000 Then
                                            XTipoPro = "TA"
                                                Else
                                            XTipoPro = "PT"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                        If Left$(Articulo, 2) = "YQ" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YH" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YP" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YF" Then
                            XTipoPro = "FA"
                        End If
                
                        ZLinea = 0
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                
                        Select Case ZLinea
                            Case 8
                                XTipoPro = "PG"
                            Case 10, 20, 22, 24, 25, 26, 27, 28, 29, 30
                                XTipoPro = "FA"
                            Case Else
                        End Select
                
                        If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                        
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
                            
                            ZArticulo = Articulo
                            ZProducto = Articulo
                            ZLote = WWLote
                            ZCliente = Cliente.Text
                                
                            ZZEntra = "N"
                            Erase ZOpcion
                            
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                            
                            If ZZEntra = "N" Then
                                ZSql = ""
                                ZSql = ZSql & "Select *"
                                ZSql = ZSql & " FROM AltaCertificado"
                                ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                                ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                                spAltaCertificado = ZSql
                                Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstAltaCertificado.RecordCount > 0 Then
                                    ZOpcion(1) = rstAltaCertificado!Opcion1
                                    ZOpcion(2) = rstAltaCertificado!Opcion2
                                    ZOpcion(3) = rstAltaCertificado!Opcion3
                                    ZOpcion(4) = rstAltaCertificado!Opcion4
                                    ZOpcion(5) = rstAltaCertificado!Opcion5
                                    ZOpcion(6) = rstAltaCertificado!Opcion6
                                    ZOpcion(7) = rstAltaCertificado!Opcion7
                                    ZOpcion(8) = rstAltaCertificado!Opcion8
                                    ZOpcion(9) = rstAltaCertificado!Opcion9
                                    ZOpcion(10) = rstAltaCertificado!Opcion10
                                    rstAltaCertificado.Close
                                    ZZEntra = "S"
                                End If
                            End If
                            
                            Call Conecta_Empresa
                                
                            If ZZEntra = "S" Then
                                If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                    ZZEntra = "N"
                                End If
                            End If
                            
                            If ZZEntra = "N" Then
                                m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                                Aaa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                ZAprueba = "N"
                            End If
                                            
                        End If
                    
                            Else
                            
                        If Left$(Articulo, 2) = "DY" Then
                            
                            ZZPartiOri = Trim(WWLote)
                            ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                            ZZEstado = Dir(ZZRuta)
                            ZZEstado = Trim(ZZEstado)
                            If ZZEstado = "" Then
                                m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                ssa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                ZAprueba = "N"
                            End If
                            
                        End If
                        
                    End If
                End If
                    
            Next ZZCiclo
        End If
        
    Next A

End Sub











Private Sub Command1_Click()

    ZZNumeroHoja = LoteImpre.Text
    Call Envio_Hoja_Email
    
    Listado.WindowTitle = "Impresion de Hoja de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
 
    Listado.ReportFileName = "HojaProduccion.rpt"
    
    Listado.GroupSelectionFormula = "{ImpreHoja44000.Hoja} in " + ZZNumeroHoja + " to " + ZZNumeroHoja
    Listado.SelectionFormula = "{ImpreHoja44000.Hoja} in " + ZZNumeroHoja + " to " + ZZNumeroHoja

    Listado.SQLQuery = "SELECT ImpreHoja44000.Hoja, ImpreHoja44000.Renglon, ImpreHoja44000.Fecha, ImpreHoja44000.Codigo1, ImpreHoja44000.Codigo2, ImpreHoja44000.Articulo1, ImpreHoja44000.Articulo2, ImpreHoja44000.Descripcion, ImpreHoja44000.Canti1, ImpreHoja44000.Lote1, ImpreHoja44000.Canti2, ImpreHoja44000.Lote2, ImpreHoja44000.Canti3, ImpreHoja44000.Lote3, ImpreHoja44000.Teorico, ImpreHoja44000.CantidadReal, ImpreHoja44000.VersionI, ImpreHoja44000.VersionII, ImpreHoja44000.VersionIII, ImpreHoja44000.LoteOri1, ImpreHoja44000.LoteOri2, ImpreHoja44000.LoteOri3, ImpreHoja44000.Nombre " _
        + "From " _
        + DSQ + ".dbo.ImpreHoja44000 ImpreHoja44000 " _
        + "WHERE " _
        + "ImpreHoja44000.Hoja >= " + ZZNumeroHoja + " AND " _
        + "ImpreHoja44000.Hoja <= " + ZZNumeroHoja
        
    Listado.EMailToList = "CAROLINA.PENZ@TANATEXCHEMICALS.COM; drodriguez@surfactan.com.ar"
 Rem   Listado.EMailToList = "hsfrias@yahoo.com.ar"
    Listado.EMailSubject = "Hojas de Produccion de SURFACTAN S.A."
    Listado.EMailMessage = "Les envio las hojas de produccion del material a remitir"
    
    Listado.Destination = crptMapi
    Listado.PrintFileName = "Hojas.doc"
    Listado.PrintFileType = crptWinWord

    MiRuta = CurDir + "\"
    MiRutaII = Left$(CurDir, 1)

    Listado.Connect = Connect()
    Listado.Action = 1
                    
    ChDrive MiRutaII
    ChDir MiRuta
    
End Sub


Sub Envio_Hoja_Email()
        
    WHoja = ZZNumeroHoja
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpreHoja44000"
    ZSql = ZSql + " Where ImpreHoja44000.Hoja = " + "'" + WHoja + "'"
    spImpreHoja = ZSql
    Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
    
  
    spHoja = "ListaHoja " + "'" + WHoja + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WFecha = rstHoja!Fecha
        WCantidadReal = Str$(rstHoja!Real)
        WTeorico = Str$(rstHoja!Teorico)
        WProducto = rstHoja!Producto
        WVersionI = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
        WVersionII = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
        WVersionIII = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
        rstHoja.Close
    End If
    
    WNombre = ""
    spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WNombre = rstTerminado!Descripcion
        rstTerminado.Close
    End If
    
    
    WCodigo1 = Left$(WProducto, 2)
    WCodigo2 = Mid$(WProducto, 4, 5) + "/" + Right$(WProducto, 3)
    
    ZZRenglon = 0
    Erase ZZGrilla
    
    spHoja = "ListaHoja " + "'" + WHoja + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZZRenglon = ZZRenglon + 1
                
                    ZZGrilla(ZZRenglon, 1) = rstHoja!Tipo
                    ZZGrilla(ZZRenglon, 2) = rstHoja!Terminado
                    ZZGrilla(ZZRenglon, 3) = rstHoja!Articulo
                    ZZGrilla(ZZRenglon, 5) = Pusing("###,###.##", rstHoja!Cantidad)
                    ZZGrilla(ZZRenglon, 6) = Str$(rstHoja!Canti1)
                    ZZGrilla(ZZRenglon, 7) = Str$(rstHoja!lote1)
                    ZZGrilla(ZZRenglon, 8) = Str$(rstHoja!Canti2)
                    ZZGrilla(ZZRenglon, 9) = Str$(rstHoja!lote2)
                    ZZGrilla(ZZRenglon, 10) = Str$(rstHoja!Canti3)
                    ZZGrilla(ZZRenglon, 11) = Str$(rstHoja!lote3)
                    
                    ZSuma = rstHoja!Canti1 + rstHoja!Canti2 + rstHoja!Canti3
                    If ZSuma = 0 Then
                        ZZGrilla(ZZRenglon, 6) = Str$(rstHoja!Cantidad)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For Da = 1 To ZZRenglon
            
        Tipo = ZZGrilla(Da, 1)
        Auxi1 = ZZGrilla(Da, 2)
        Auxi2 = ZZGrilla(Da, 3)
                
        Select Case Tipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZGrilla(Da, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZGrilla(Da, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            Case Else
        End Select
    Next Da

    ZZLugar = 0

    For A = 1 To 40
    
        WTipo = UCase(ZZGrilla(A, 1))
        
        If Trim(WTipo) <> "" Then
        
            ZZLugar = ZZLugar + 1
            
            WTerminado = UCase(ZZGrilla(A, 2))
            WArticulo = UCase(ZZGrilla(A, 3))
            WDescripcion = UCase(ZZGrilla(A, 4))
            WCantidad = ZZGrilla(A, 5)
            WLinea = Str$(A)
            
            If WTipo = "M" Then
                WArticulo1 = Left$(WArticulo, 2)
                WArticulo2 = Mid$(WArticulo, 4, 3) + "-" + Right$(WArticulo, 3)
                    Else
                WArticulo1 = Left$(WTerminado, 2)
                WArticulo2 = Mid$(WTerminado, 4, 5) + "-" + Right$(WTerminado, 3)
            End If
            
            WCanti1 = ZZGrilla(A, 6)
            WLote1 = ZZGrilla(A, 7)
            WLoteOri1 = ""
            WCanti2 = ZZGrilla(A, 8)
            WLote2 = ZZGrilla(A, 9)
            WLoteOri2 = ""
            WCanti3 = ZZGrilla(A, 10)
            WLote3 = ZZGrilla(A, 11)
            WLoteOri3 = ""
        
            For ZZPasalote = 1 To 3
                
                Select Case ZZPasalote
                    Case 1
                        XXLote = ZZGrilla(A, 7)
                    Case 2
                        XXLote = ZZGrilla(A, 9)
                    Case Else
                        XXLote = ZZGrilla(A, 11)
                End Select
         
                If WTipo = "M" Then
                    XParam = "'" + XXLote + "','" _
                                + WArticulo + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        Select Case ZZPasalote
                            Case 1
                                WLoteOri1 = rstLaudo!partiori
                            Case 2
                                WLoteOri2 = rstLaudo!partiori
                            Case Else
                                WLoteOri3 = rstLaudo!partiori
                        End Select
                        rstLaudo.Close
                    End If
                End If
                
            Next ZZPasalote
                
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ImpreHoja44000 ("
            ZSql = ZSql & "Hoja ,"
            ZSql = ZSql & "Renglon ,"
            ZSql = ZSql & "Fecha ,"
            ZSql = ZSql & "Codigo1 ,"
            ZSql = ZSql & "Codigo2 ,"
            ZSql = ZSql & "Nombre ,"
            ZSql = ZSql & "Articulo1 ,"
            ZSql = ZSql & "Articulo2 ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Canti1 ,"
            ZSql = ZSql & "Lote1 ,"
            ZSql = ZSql & "LoteOri1 ,"
            ZSql = ZSql & "Canti2 ,"
            ZSql = ZSql & "Lote2 ,"
            ZSql = ZSql & "LoteOri2 ,"
            ZSql = ZSql & "Canti3 ,"
            ZSql = ZSql & "Lote3 ,"
            ZSql = ZSql & "LoteOri3 ,"
            ZSql = ZSql & "Teorico ,"
            ZSql = ZSql & "CantidadReal ,"
            ZSql = ZSql & "VersionI ,"
            ZSql = ZSql & "VersionII ,"
            ZSql = ZSql & "VersionIII )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WHoja + "',"
            ZSql = ZSql & "'" + WLinea + "',"
            ZSql = ZSql & "'" + WFecha + "',"
            ZSql = ZSql & "'" + WCodigo1 + "',"
            ZSql = ZSql & "'" + WCodigo2 + "',"
            ZSql = ZSql & "'" + WNombre + "',"
            ZSql = ZSql & "'" + WArticulo1 + "',"
            ZSql = ZSql & "'" + WArticulo2 + "',"
            ZSql = ZSql & "'" + WCantidad + "',"
            ZSql = ZSql & "'" + WDescripcion + "',"
            ZSql = ZSql & "'" + WCanti1 + "',"
            ZSql = ZSql & "'" + WLote1 + "',"
            ZSql = ZSql & "'" + WLoteOri1 + "',"
            ZSql = ZSql & "'" + WCanti2 + "',"
            ZSql = ZSql & "'" + WLote2 + "',"
            ZSql = ZSql & "'" + WLoteOri2 + "',"
            ZSql = ZSql & "'" + WCanti3 + "',"
            ZSql = ZSql & "'" + WLote3 + "',"
            ZSql = ZSql & "'" + WLoteOri3 + "',"
            ZSql = ZSql & "'" + WTeorico + "',"
            ZSql = ZSql & "'" + WCantidadReal + "',"
            ZSql = ZSql & "'" + ZVersionI + "',"
            ZSql = ZSql & "'" + ZVersionII + "',"
            ZSql = ZSql & "'" + ZVersionIII + "')"
    
            spImpreHoja = ZSql
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next A
                    
            
    XLinea = ZZLugar
    For Ciclo = XLinea To 14
    
        ZZLugar = ZZLugar + 1
        WLinea = Str$(ZZLugar)
                
        WArticulo1 = ""
        WArticulo2 = ""
        WCantidad = ""
        WCanti1 = ""
        WLote1 = ""
        WLoteOri1 = ""
        WCanti2 = ""
        WLote2 = ""
        WLoteOri2 = ""
        WCanti3 = ""
        WLote3 = ""
        WLoteOri3 = ""
        WDescripcion = ""
    
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreHoja44000 ("
        ZSql = ZSql & "Hoja ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Fecha ,"
        ZSql = ZSql & "Codigo1 ,"
        ZSql = ZSql & "Codigo2 ,"
        ZSql = ZSql & "Nombre ,"
        ZSql = ZSql & "Articulo1 ,"
        ZSql = ZSql & "Articulo2 ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "Descripcion ,"
        ZSql = ZSql & "Canti1 ,"
        ZSql = ZSql & "Lote1 ,"
        ZSql = ZSql & "LoteOri1 ,"
        ZSql = ZSql & "Canti2 ,"
        ZSql = ZSql & "Lote2 ,"
        ZSql = ZSql & "LoteOri2 ,"
        ZSql = ZSql & "Canti3 ,"
        ZSql = ZSql & "Lote3 ,"
        ZSql = ZSql & "LoteOri3 ,"
        ZSql = ZSql & "Teorico ,"
        ZSql = ZSql & "CantidadReal ,"
        ZSql = ZSql & "VersionI ,"
        ZSql = ZSql & "VersionII ,"
        ZSql = ZSql & "VersionIII )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + WHoja + "',"
        ZSql = ZSql & "'" + WLinea + "',"
        ZSql = ZSql & "'" + WFecha + "',"
        ZSql = ZSql & "'" + WCodigo1 + "',"
        ZSql = ZSql & "'" + WCodigo2 + "',"
        ZSql = ZSql & "'" + WNombre + "',"
        ZSql = ZSql & "'" + WArticulo1 + "',"
        ZSql = ZSql & "'" + WArticulo2 + "',"
        ZSql = ZSql & "'" + WCantidad + "',"
        ZSql = ZSql & "'" + WDescripcion + "',"
        ZSql = ZSql & "'" + WCanti1 + "',"
        ZSql = ZSql & "'" + WLote1 + "',"
        ZSql = ZSql & "'" + WLoteOri1 + "',"
        ZSql = ZSql & "'" + WCanti2 + "',"
        ZSql = ZSql & "'" + WLote2 + "',"
        ZSql = ZSql & "'" + WLoteOri2 + "',"
        ZSql = ZSql & "'" + WCanti3 + "',"
        ZSql = ZSql & "'" + WLote3 + "',"
        ZSql = ZSql & "'" + WLoteOri3 + "',"
        ZSql = ZSql & "'" + WTeorico + "',"
        ZSql = ZSql & "'" + WCantidadReal + "',"
        ZSql = ZSql & "'" + ZVersionI + "',"
        ZSql = ZSql & "'" + ZVersionII + "',"
        ZSql = ZSql & "'" + ZVersionIII + "')"

        spImpreHoja = ZSql
        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

    Next Ciclo
    

End Sub

Private Sub Graba_Click()
    WProceso = 0
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub GrabaII_Click()

    XEmpresa = Wempresa
    Wempresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)



    ZSql = ""
    ZSql = ZSql & "UPDATE Pedido SET "
    ZSql = ZSql & " MarcaFactura = " + "'" + "1" + "'"
    ZSql = ZSql & " Where Pedido = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    Call Conecta_Empresa
    
    Call cmdClose_Click
        
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WClave.Text = "DT24JF" Then
            Pass.Visible = False
            Call GrabaII_Click
        End If
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub


Private Sub Limpia_Click()

    Erase XEnvase

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    MarcaFactura.ListIndex = 0
    
    Call Limpia_Vector
    
    Renglon = 0
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    
    Descri1.Caption = ""
    Descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    Canti4.Text = ""
    Canti5.Text = ""
    
    Pedido.SetFocus

End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    WEnvase(1) = 20
    WEnvase(2) = 21
    WEnvase(3) = 22
    WEnvase(4) = 23
    WEnvase(5) = 24
    WEnvase(6) = 25
    WEnvase(7) = 26
    WEnvase(8) = 30
    WEnvase(9) = 28
 
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

    For Cicla = 1 To 9
        spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
            rstEnvase.Close
                Else
            WImpre(Cicla) = ""
        End If
    Next Cicla
    
    Call Conecta_Empresa

    Erase XEnvase
    Erase XLote
    
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    MarcaFactura.Clear
    MarcaFactura.AddItem ""
    MarcaFactura.AddItem "Disponible"
    MarcaFactura.ListIndex = 0
    
    Renglon = 0
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    ZZPasaProcesoActualiza = ""
    
    
     
End Sub

Private Sub Proceso_Click()

    Erase XEnvase
    
    Call Limpia_Vector
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    Erase ClavePedido
    Erase ZVector
    Erase XLote
    
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
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Rem dada
                    Rem ojo
                    Rem ver
                    Canti = !Cantidad - !Facturado
                    Rem Canti = !Cantidad
                    
                    If Canti > 0 Then
                
                        Rem dada
                        Rem dada
                        Rem dada
                        Rem dada
                        Rem dada
                
                        Renglon = Renglon + 1
                
                        WVector1.TextMatrix(Renglon, 1) = !Terminado
                        
                        WVector1.TextMatrix(Renglon, 3) = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                        Rem WVector1.TextMatrix(Renglon, 3) = Pusing("###,###.##", Str$(!Cantidad))
                        
                        Cantidad = IIf(IsNull(rstPedido!Cantidad1), "0", rstPedido!Cantidad1)
                        WVector1.TextMatrix(Renglon, 4) = Pusing("###,###.##", Str$(Cantidad))
                
                        Rem Resta = IIf(IsNull(rstPedido!Cantidad2), "0", rstPedido!Cantidad2)
                        If rstPedido!lote1 <> 0 And rstPedido!lote2 <> 0 Then
                            WVector1.TextMatrix(Renglon, 5) = "Varios"
                                Else
                            WVector1.TextMatrix(Renglon, 5) = rstPedido!lote1
                        End If
                                    
                        Auxi1 = !Terminado
                    
                        If Resta <> 0 Or Cantidad <> 0 Then
                            WVector1.TextMatrix(Renglon, 6) = "S"
                        End If
                    
                        WLugar = Renglon
                        
                        
                        
                        XLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        XLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        XLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        XLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        XLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        XLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        XLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        XLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        XLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        XLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                        XLote(WLugar, 11) = IIf(IsNull(rstPedido!lote6), "0", rstPedido!lote6)
                        XLote(WLugar, 12) = IIf(IsNull(rstPedido!CantiLote6), "0", rstPedido!CantiLote6)
                        XLote(WLugar, 13) = IIf(IsNull(rstPedido!lote7), "0", rstPedido!lote7)
                        XLote(WLugar, 14) = IIf(IsNull(rstPedido!CantiLote7), "0", rstPedido!CantiLote7)
                        XLote(WLugar, 15) = IIf(IsNull(rstPedido!lote8), "0", rstPedido!lote8)
                        XLote(WLugar, 16) = IIf(IsNull(rstPedido!CantiLote8), "0", rstPedido!CantiLote8)
                        XLote(WLugar, 17) = IIf(IsNull(rstPedido!lote9), "0", rstPedido!lote9)
                        XLote(WLugar, 18) = IIf(IsNull(rstPedido!CantiLote9), "0", rstPedido!CantiLote9)
                        XLote(WLugar, 19) = IIf(IsNull(rstPedido!lote10), "0", rstPedido!lote10)
                        XLote(WLugar, 20) = IIf(IsNull(rstPedido!CantiLote10), "0", rstPedido!CantiLote10)
                        XLote(WLugar, 21) = IIf(IsNull(rstPedido!lote11), "0", rstPedido!lote11)
                        XLote(WLugar, 22) = IIf(IsNull(rstPedido!CantiLote11), "0", rstPedido!CantiLote11)
                        XLote(WLugar, 23) = IIf(IsNull(rstPedido!lote12), "0", rstPedido!lote12)
                        XLote(WLugar, 24) = IIf(IsNull(rstPedido!CantiLote12), "0", rstPedido!CantiLote12)
                    
                    
                        ZZEnvase = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 31) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                            XLote(WLugar, 32) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env2), "0", rstPedido!Env2)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv2), "0", rstPedido!CantiEnv2)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 33) = IIf(IsNull(rstPedido!Env2), "0", rstPedido!Env2)
                            XLote(WLugar, 34) = IIf(IsNull(rstPedido!CantiEnv2), "0", rstPedido!CantiEnv2)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                            
                        ZZEnvase = IIf(IsNull(rstPedido!Env3), "0", rstPedido!Env3)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv3), "0", rstPedido!CantiEnv3)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 35) = IIf(IsNull(rstPedido!Env3), "0", rstPedido!Env3)
                            XLote(WLugar, 36) = IIf(IsNull(rstPedido!CantiEnv3), "0", rstPedido!CantiEnv3)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env4), "0", rstPedido!Env4)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv4), "0", rstPedido!CantiEnv4)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 37) = IIf(IsNull(rstPedido!Env4), "0", rstPedido!Env4)
                            XLote(WLugar, 38) = IIf(IsNull(rstPedido!CantiEnv4), "0", rstPedido!CantiEnv4)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env5), "0", rstPedido!Env5)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv5), "0", rstPedido!CantiEnv5)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 39) = IIf(IsNull(rstPedido!Env5), "0", rstPedido!Env5)
                            XLote(WLugar, 40) = IIf(IsNull(rstPedido!CantiEnv5), "0", rstPedido!CantiEnv5)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env6), "0", rstPedido!Env6)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv6), "0", rstPedido!CantiEnv6)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 41) = IIf(IsNull(rstPedido!Env6), "0", rstPedido!Env6)
                            XLote(WLugar, 42) = IIf(IsNull(rstPedido!CantiEnv6), "0", rstPedido!CantiEnv6)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env7), "0", rstPedido!Env7)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv7), "0", rstPedido!CantiEnv7)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 43) = IIf(IsNull(rstPedido!Env7), "0", rstPedido!Env7)
                            XLote(WLugar, 44) = IIf(IsNull(rstPedido!CantiEnv7), "0", rstPedido!CantiEnv7)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env8), "0", rstPedido!Env8)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv8), "0", rstPedido!CantiEnv8)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 45) = IIf(IsNull(rstPedido!Env8), "0", rstPedido!Env8)
                            XLote(WLugar, 46) = IIf(IsNull(rstPedido!CantiEnv8), "0", rstPedido!CantiEnv8)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env9), "0", rstPedido!Env9)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv9), "0", rstPedido!CantiEnv9)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 47) = IIf(IsNull(rstPedido!Env9), "0", rstPedido!Env9)
                            XLote(WLugar, 48) = IIf(IsNull(rstPedido!CantiEnv9), "0", rstPedido!CantiEnv9)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env10), "0", rstPedido!Env10)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv10), "0", rstPedido!CantiEnv10)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 49) = IIf(IsNull(rstPedido!Env10), "0", rstPedido!Env10)
                            XLote(WLugar, 50) = IIf(IsNull(rstPedido!CantiEnv10), "0", rstPedido!CantiEnv10)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env11), "0", rstPedido!Env11)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv11), "0", rstPedido!CantiEnv11)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 51) = IIf(IsNull(rstPedido!Env11), "0", rstPedido!Env11)
                            XLote(WLugar, 52) = IIf(IsNull(rstPedido!CantiEnv11), "0", rstPedido!CantiEnv11)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        ZZEnvase = IIf(IsNull(rstPedido!Env12), "0", rstPedido!Env12)
                        ZZCantiEnvase = IIf(IsNull(rstPedido!CantiEnv12), "0", rstPedido!CantiEnv12)
                        If ZZEnvase <> 31 Then
                            XLote(WLugar, 53) = IIf(IsNull(rstPedido!Env12), "0", rstPedido!Env12)
                            XLote(WLugar, 54) = IIf(IsNull(rstPedido!CantiEnv12), "0", rstPedido!CantiEnv12)
                                Else
                            XLote(WLugar, 81) = Str$(Val(XLote(WLugar, 81)) + ZZCantiEnvase)
                        End If
                        
                        XLote(WLugar, 61) = IIf(IsNull(rstPedido!Bultos1), "0", rstPedido!Bultos1)
                        XLote(WLugar, 62) = IIf(IsNull(rstPedido!Bultos2), "0", rstPedido!Bultos2)
                        XLote(WLugar, 63) = IIf(IsNull(rstPedido!Bultos3), "0", rstPedido!Bultos3)
                        XLote(WLugar, 64) = IIf(IsNull(rstPedido!Bultos4), "0", rstPedido!Bultos4)
                        XLote(WLugar, 65) = IIf(IsNull(rstPedido!Bultos5), "0", rstPedido!Bultos5)
                        XLote(WLugar, 66) = IIf(IsNull(rstPedido!Bultos6), "0", rstPedido!Bultos6)
                        XLote(WLugar, 67) = IIf(IsNull(rstPedido!Bultos7), "0", rstPedido!Bultos7)
                        XLote(WLugar, 68) = IIf(IsNull(rstPedido!Bultos8), "0", rstPedido!Bultos8)
                        XLote(WLugar, 69) = IIf(IsNull(rstPedido!Bultos9), "0", rstPedido!Bultos9)
                        XLote(WLugar, 70) = IIf(IsNull(rstPedido!Bultos10), "0", rstPedido!Bultos10)
                        XLote(WLugar, 71) = IIf(IsNull(rstPedido!Bultos11), "0", rstPedido!Bultos11)
                        XLote(WLugar, 72) = IIf(IsNull(rstPedido!Bultos12), "0", rstPedido!Bultos12)
                    
                        Auxiliar(Renglon, 1) = Auxi1
                        Auxiliar(Renglon, 2) = Canti
                        
                        ClavePedido(Renglon) = rstPedido!Clave
                    
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        ZVector(Renglon, 1) = !Terminado
                        ZVector(Renglon, 2) = ""
                        ZVector(Renglon, 3) = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                        ZVector(Renglon, 4) = ""
                        ZVector(Renglon, 5) = IIf(IsNull(rstPedido!Especificaciones), "0", rstPedido!Especificaciones)
                        
                    End If
        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri1.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
    
    Call Conecta_Empresa
    
    If Val(Envase1.Text) = 0 Then
        Envase1.Text = ""
    End If
    If Val(Envase2.Text) = 0 Then
        Envase2.Text = ""
    End If
    If Val(Envase3.Text) = 0 Then
        Envase3.Text = ""
    End If
    If Val(Envase4.Text) = 0 Then
        Envase4.Text = ""
    End If
    If Val(Envase5.Text) = 0 Then
        Envase5.Text = ""
    End If
    
    If Val(Canti1.Text) = 0 Then
        Canti1.Text = ""
    End If
    If Val(Canti2.Text) = 0 Then
        Canti2.Text = ""
    End If
    If Val(Canti3.Text) = 0 Then
        Canti3.Text = ""
    End If
    If Val(Canti4.Text) = 0 Then
        Canti4.Text = ""
    End If
    If Val(Canti5.Text) = 0 Then
        Canti5.Text = ""
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(Da, 1)
        Canti = Auxiliar(Da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) <> "PT" And Left$(Auxi1, 2) <> "YQ" And Left$(Auxi1, 2) <> "YF" And Left$(Auxi1, 2) <> "YP" And Left$(Auxi1, 2) <> "YH" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                ClavePreciosMp = Cliente.Text + Auxi1
                
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
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Precio = rstPreciosMp!Precio
                    ZVector(Renglon, 4) = Str$(Precio)
                    rstPreciosMp.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    ZVector(Renglon, 2) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 23 Step 2
                
                    If Val(XLote(Da, Ciclo)) = 0 Then
                    
                        XLote(Da, Ciclo) = ""
                        
                            Else
                            
                        ZEntra = "N"
                        
                        XEmpresa = Wempresa
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    Wempresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    Wempresa = "0011"
                                    txtOdbc = "Empresa11"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    Wempresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + XLote(Da, Ciclo) + "'"
                        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArti + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            XLote(Da, Ciclo) = IIf(IsNull(rstLaudo!partiori), "", rstLaudo!partiori)
                            ZEntra = "S"
                            rstLaudo.Close
                        End If
                        
                        If ZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Lote = " + "'" + XLote(Da, Ciclo) + "'"
                            ZSql = ZSql + " and Guia.Articulo = " + "'" + WArti + "'"
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                XLote(Da, Ciclo) = IIf(IsNull(rstMovguia!partiori), "", rstMovguia!partiori)
                                ZEntra = "S"
                                rstMovguia.Close
                            End If
                        End If
                        
                        Call Conecta_Empresa
                            
                        Rem XParam = "'" + xLote(Da, Ciclo) + "'"
                        Rem spLaudo = "ListaLaudo " + XParam
                        Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstLaudo.RecordCount > 0 Then
                        Rem     xLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                        Rem     rstLaudo.Close
                        Rem End If
                        
                    End If
                Next Ciclo

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
            
            Case Else
            
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
                
            
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstPrecios!Descripcion
                    Precio = rstPrecios!Precio
                    ZVector(Renglon, 2) = rstPrecios!Descripcion
                    ZVector(Renglon, 4) = Str$(Precio)
                    rstPrecios.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 23 Step 2
                    If Val(XLote(Da, Ciclo)) = 0 Then
                        XLote(Da, Ciclo) = ""
                    End If
                Next Ciclo

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
                
        End Select
        
    Next Da
    
    WVector1.TopRow = 1
    WVector1.Row = 1
    WVector1.Col = 1
    
    Call StartEdit
    
    Graba.Enabled = True

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
                
            WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
            WMarcaFactura = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
            Cliente.Text = rstPedido!Cliente
            Fecha.Text = rstPedido!Fecha
            WFecEntrega = rstPedido!FecEntrega
            WObservaciones = rstPedido!Observaciones
            ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
            
            Select Case rstPedido!TipoPedido
                Case 1
                    WTipoPedido = "CO"
                Case 3
                    WTipoPedido = "BI"
                Case 4
                    WTipoPedido = "FA"
                Case 5
                    WTipoPedido = "PG"
                Case Else
                    WTipoPedido = "PT"
            End Select
            
            If Left$(rstPedido!Terminado, 4) = "PT-4" Then
                WTipoPedido = "TA"
            End If
            
            rstPedido.Close
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                
                Rem WDirentrega = rstCliente!DirEntrega
                WDirentrega = ""
    
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
    
                WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                
                WPago = Str$(rstCliente!Pago1)
                rstCliente.Close
                spPago = "ConsultaPago " + "'" + WPago + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstPago.RecordCount > 0 Then
                    WDespago = rstPago!Nombre
                    rstPago.Close
                End If
            End If
            
            Call Conecta_Empresa
            
            Call Proceso_Click
            
            WVector1.TopRow = 1
            WVector1.Row = 1
            WVector1.Col = 4
            Call StartEdit
                
                Else
            
            Call Conecta_Empresa
            
        End If
        
    End If
End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
        
        Call Conecta_Empresa
        
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Envase5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
    
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    If WLote1.Text <> "" Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If WLote2.Text <> "" Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If WLote3.Text <> "" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    If WLote4.Text <> "" Then
        Suma = Suma + Val(WCanti4.Text)
    End If
    If WLote5.Text <> "" Then
        Suma = Suma + Val(WCanti5.Text)
    End If
    
    Rem by nan
    If Suma = XCantidad Then
        WEstado = "S"
            Else
        m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
   End If
    
    If WEstado = "S" Then
    
        Erase ControlLote
        ControlLote(1, 1) = WLote1.Text
        ControlLote(1, 2) = WCanti1.Text
        ControlLote(2, 1) = WLote2.Text
        ControlLote(2, 2) = WCanti2.Text
        ControlLote(3, 1) = WLote3.Text
        ControlLote(3, 2) = WCanti3.Text
        ControlLote(4, 1) = WLote4.Text
        ControlLote(4, 2) = WCanti4.Text
        ControlLote(5, 1) = WLote5.Text
        ControlLote(5, 2) = WCanti5.Text
        ControlLote(6, 1) = WLote6.Text
        ControlLote(6, 2) = WCanti6.Text
        ControlLote(7, 1) = WLote7.Text
        ControlLote(7, 2) = WCanti7.Text
        ControlLote(8, 1) = WLote8.Text
        ControlLote(8, 2) = WCanti8.Text
        ControlLote(9, 1) = WLote9.Text
        ControlLote(9, 2) = WCanti9.Text
        ControlLote(10, 1) = WLote10.Text
        ControlLote(10, 2) = WCanti10.Text
        ControlLote(11, 1) = WLote11.Text
        ControlLote(11, 2) = WCanti11.Text
        ControlLote(12, 1) = WLote12.Text
        ControlLote(12, 2) = WCanti12.Text
    
        For Ciclo1 = 1 To 12
            If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                For Ciclo2 = 1 To 5
                    If Ciclo1 <> Ciclo2 Then
                        If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                            m$ = "A asignado una misma partida 2 veces"
                            A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                            Exit For
                        End If
                    End If
                Next Ciclo2
            End If
            If WEstado = "N" Then
                Exit For
            End If
        Next Ciclo1
        
    End If

    If WEstado = "S" Then
    
        Erase ControlLote
        ControlLote(1, 1) = WLote1.Text
        ControlLote(1, 2) = WCanti1.Text
        ControlLote(2, 1) = WLote2.Text
        ControlLote(2, 2) = WCanti2.Text
        ControlLote(3, 1) = WLote3.Text
        ControlLote(3, 2) = WCanti3.Text
        ControlLote(4, 1) = WLote4.Text
        ControlLote(4, 2) = WCanti4.Text
        ControlLote(5, 1) = WLote5.Text
        ControlLote(5, 2) = WCanti5.Text
        ControlLote(6, 1) = WLote6.Text
        ControlLote(6, 2) = WCanti6.Text
        ControlLote(7, 1) = WLote7.Text
        ControlLote(7, 2) = WCanti7.Text
        ControlLote(8, 1) = WLote8.Text
        ControlLote(8, 2) = WCanti8.Text
        ControlLote(9, 1) = WLote9.Text
        ControlLote(9, 2) = WCanti9.Text
        ControlLote(10, 1) = WLote10.Text
        ControlLote(10, 2) = WCanti10.Text
        ControlLote(11, 1) = WLote11.Text
        ControlLote(11, 2) = WCanti11.Text
        ControlLote(12, 1) = WLote12.Text
        ControlLote(12, 2) = WCanti12.Text
    
        For Ciclo1 = 1 To 12
    
            WLote = ControlLote(Ciclo1, 1)
            WCanti = Val(ControlLote(Ciclo1, 2))
            
            If WLote <> "" Or Val(WCanti) <> 0 Then
            
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                Wempresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "TA"
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                    
                    ZSql = ""
                    If Val(WLote) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            Call Redondeo(WSaldo)
                            WEntra = "S"
                            If WSaldo < WCanti Then
                                m$ = "La cantidad informada supera al saldo disponible"
                                A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                WEstado = "N"
                            End If
                            ZEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            ZEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                            If ZEstado = "N" Then
                                If ZEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WEstado = "N"
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        ZSql = ""
                        If Val(WLote) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstMovguia.Close
                            End With
                        End If
                    End If
                    
                    Call Conecta_Empresa
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
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
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = Wempresa
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    Wempresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    Wempresa = "0011"
                                    txtOdbc = "Empresa11"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "TA"
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    Wempresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
                        XParam = "'" + WLote + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            Call Redondeo(WSaldo)
                            WEntra = "S"
                            If WSaldo < WCanti Then
                                m$ = "La cantidad informada supera al saldo disponible"
                                A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                WEstado = "N"
                            End If
                            ZEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                            ZEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                            If ZEstado = "N" Then
                                If ZEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WEstado = "N"
                            End If
                            WFechaHoja = rstHoja!Fecha
                            rstHoja.Close
                            Rem WVida = 0
                            Rem
                            Rem spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                            Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            Rem If rstTerminado.RecordCount > 0 Then
                            Rem     WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                            Rem     rstTerminado.Close
                            Rem End If
                            Rem
                            Rem If WVida <> 0 Then
                            Rem
                            Rem     WMes = Val(Mid$(WFechaHoja, 4, 2))
                            Rem     WAno = Val(Right$(WFechaHoja, 4))
                            Rem     For Ciclo = 1 To WVida
                            Rem         WMes = WMes + 1
                            Rem         If WMes > 12 Then
                            Rem             WAno = WAno + 1
                            Rem             WMes = 1
                            Rem         End If
                            Rem     Next Ciclo
                            Rem     XMes = Str$(WMes)
                            Rem     XAno = Str$(WAno)
                            Rem     Call Ceros(XMes, 2)
                            Rem     Call Ceros(XAno, 4)
                            Rem     WVencimiento = "01/" + XMes + "/" + XAno
                            Rem
                            Rem     WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            Rem     WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
                            Rem
                            Rem     WFechaVencimiento = "01" + Mid$(WVencimiento, 3, 10)
                            Rem     WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
                            Rem
                            Rem     Pasa = "S"
                            Rem     If WFechaActualOrd >= WFechaVencimientoOrd Then
                            Rem         Pasa = "N"
                            Rem             Else
                            Rem         Meses = 0
                            Rem         WMes = Val(Mid$(WFechaActual, 4, 2))
                            Rem         WAno = Val(Right$(WFechaActual, 4))
                            Rem         Do
                            Rem             Meses = Meses + 1
                            Rem             WMes = WMes + 1
                            Rem             If WMes > 12 Then
                            Rem                 WAno = WAno + 1
                            Rem                 WMes = 1
                            Rem             End If
                            Rem             XMes = Str$(WMes)
                            Rem             XAno = Str$(WAno)
                            Rem             Call Ceros(XMes, 2)
                            Rem             Call Ceros(XAno, 4)
                            Rem             WCompara = "01/" + XMes + "/" + XAno
                            Rem             If WCompara = WFechaVencimiento Then
                            Rem                 Exit Do
                            Rem             End If
                            Rem         Loop
                            Rem         If Meses <= 12 Then
                            Rem             Pasa = "N"
                            Rem         End If
                            Rem     End If
                            Rem
                            Rem     If Pasa = "N" Then
                            Rem         m$ = "EL Producto tiene menos de un año de vida util"
                            Rem         G% = MsgBox(m$, 0, "Actualizacion de Pedido")
                            Rem         WEstado = "N"
                            Rem     End If
                            Rem
                            Rem End If
                        End If
                
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                        + WLote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstMovguia.Close
                            End If
                        End If
                                
                        Call Conecta_Empresa
                        
                
                            Else
                            
                        WEntra = "S"
                        
                    End If
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
            End Select
            
            End If
            
        Next Ciclo1

    End If
    
End Sub

Private Sub reImpre_Click()
    Call Impresion
End Sub

Private Sub Impresion()

    On Error GoTo WError
    
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
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        ZZVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
        ZZMarcaFactura = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
        ZZCliente = rstPedido!Cliente
        ZZFecha = rstPedido!Fecha
        ZZFecEntrega = rstPedido!FecEntrega
        ZZObservaciones = rstPedido!Observaciones
        ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
        ZZTipoped = rstPedido!Tipoped
        ZZVia = rstPedido!Via
        ZZOrden = rstPedido!OrdenCpa
        
        Select Case rstPedido!TipoPedido
            Case 1
                ZZTipoPedido = "CO"
            Case 3
                ZZTipoPedido = "BI"
            Case 4
                ZZTipoPedido = "FA"
            Case 5
                ZZTipoPedido = "PG"
            Case Else
                ZZTipoPedido = "PT"
        End Select
            
        rstPedido.Close
        
        spCliente = "ConsultaCliente " + "'" + ZZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZDesCliente = rstCliente!Razon
            
            ZZDirentrega = ""
            ZZDesPago = ""

            ZDirEntrega(1) = rstCliente!DirEntrega
            ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
            ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
            ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
            ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))

            ZZDirentrega = ZDirEntrega(ZZLugarDirEntrega)
            ZZPago = Str$(rstCliente!Pago1)
            
            Erase WEspecif
            
            WEspecif(1) = ""
            WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
            WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
            WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
            WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
            WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
            WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
            WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
            WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
            WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
            For CicloEspecif = 1 To 10
                WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
            Next CicloEspecif
            
            rstCliente.Close
            
            spPago = "ConsultaPago " + "'" + ZZPago + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                ZZDesPago = rstPago!Nombre
                rstPago.Close
            End If
            
        End If
        
    End If
        
    Call Conecta_Empresa
    
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    WObservaciones = Left$(ZZObservaciones + Space$(100), 100)
    
    WTipoPedido = ""
    Select Case ZZTipoped
        Case 0
            WTipoPedido = " (Normal)"
        Case 1
            WTipoPedido = " (A fecha)"
        Case 2
            WTipoPedido = " (Fecha Limite)"
        Case 3
            WTipoPedido = " (Urgente)"
        Case 4
            WTipoPedido = " (Retira Cliente)"
        Case 5
            WTipoPedido = " (Muestra)"
        Case Else
    End Select
    
    WVia = ""
    Select Case ZZVia
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
    
    Suma = 0
    Renglon = 0
    WRenglon = 0
        
    For A = 1 To 99
        
        Suma = Suma + 1
        Renglon = Renglon + 1
        
        WLugar = Renglon
                
        ZZArticulo = ZVector(WLugar, 1)
        ZZDescripcion = ZVector(WLugar, 2)
        ZZCantidad = ZVector(WLugar, 3)
        ZZPrecio = ZVector(WLugar, 4)
        WEspecificaciones = ZVector(WLugar, 5)
            
        If Val(ZZCantidad) <> 0 Then
        
            Erase ImpreEnvase
            LugarEnvase = 0
            
            For Cicla = 1 To 6 Step 2
                If Val(XEnvase(WLugar, Cicla)) <> 0 Then
                    LugarEnvase = LugarEnvase + 1
                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WLugar, Cicla) + "'"
                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If rstEnvase.RecordCount > 0 Then
                        WAbre = rstEnvase!Abreviatura
                        rstEnvase.Close
                            Else
                        WAbre = ""
                    End If
                    
                    ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WLugar, Cicla + 1))) + " " + Left$(WAbre, 8)
                End If
            Next Cicla
            
            WRenglon = WRenglon + 1
            
            Auxi = Pedido.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            ZClave = "1" + Auxi + Auxi1
            ZTipo = "1"
            ZPedido = Pedido.Text
            ZRenglon = Str$(WRenglon)
            ZEmpresa = WNombreEmpresa
            ZVersion = ZZVersion
            ZCliente = Cliente.Text
            ZNombre = DesCliente.Caption
            ZFecha = Fecha.Text
            ZFechaent = ZZFecEntrega
            ZTipoPedido = WTipoPedido
            ZCondicion = ZZDesPago
            ZEntrega = ZZDirentrega
            ZObservaciones1 = Left$(ZZObservaciones, 50)
            ZObservaciones2 = Right$(ZZObservaciones, 50)
            ZOrden = ZZOrden
            ZArticulo = ZZArticulo
            ZDescripcion = ZZDescripcion
            ZPrecio = ZZPrecio
            ZCantidad = ZZCantidad
            ZEnvase = ImpreEnvase(1)
            
            spImprePed = "INSERT INTO ImprePed (" + _
                        "Clave ," + _
                        "Tipo , Pedido ," + _
                        "Renglon , Empresa ," + _
                        "Version , Cliente ," + _
                        "Nombre , Fecha ," + _
                        "Fechaent , TipoPedido ," + _
                        "Condicion , Entrega ," + _
                        "Observaciones1 , Observaciones2 ," + _
                        "Orden , Articulo ," + _
                        "Descripcion , Precio ," + _
                        "Cantidad , Envase )" + _
                        "Values (" + _
                        "'" + ZClave + "'," + _
                        "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                        "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                        "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                        "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                        "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                        "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                        "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                        "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                        "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
            Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
            
            If WEspecificaciones <> "" And WEspecificaciones <> "0" Then
            
                WRenglon = WRenglon + 1
                
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = Pedido.Text
                ZRenglon = Str$(WRenglon)
                ZEmpresa = WNombreEmpresa
                ZVersion = ZZVersion
                ZCliente = Cliente.Text
                ZNombre = DesCliente.Caption
                ZFecha = Fecha.Text
                ZFechaent = ZZFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = ZZDesPago
                ZEntrega = ZZDirentrega
                ZObservaciones1 = Left$(ZZObservaciones, 50)
                ZObservaciones2 = Right$(ZZObservaciones, 50)
                ZOrden = ZZOrden
                ZArticulo = "Especif.:"
                ZDescripcion = WEspecificaciones
                ZPrecio = "0"
                ZCantidad = "0"
                ZEnvase = ""
                
                spImprePed = "INSERT INTO ImprePed (" + _
                        "Clave ," + _
                        "Tipo , Pedido ," + _
                        "Renglon , Empresa ," + _
                        "Version , Cliente ," + _
                        "Nombre , Fecha ," + _
                        "Fechaent , TipoPedido ," + _
                        "Condicion , Entrega ," + _
                        "Observaciones1 , Observaciones2 ," + _
                        "Orden , Articulo ," + _
                        "Descripcion , Precio ," + _
                        "Cantidad , Envase )" + _
                        "Values (" + _
                        "'" + ZClave + "'," + _
                        "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                        "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                        "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                        "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                        "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                        "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                        "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                        "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                        "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
                Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            For Ciclo = 2 To LugarEnvase
            
                WRenglon = WRenglon + 1
                
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = Pedido.Text
                ZRenglon = Str$(WRenglon)
                ZEmpresa = WNombreEmpresa
                ZVersion = ZZVersion
                ZCliente = Cliente.Text
                ZNombre = DesCliente.Caption
                ZFecha = Fecha.Text
                ZFechaent = ZZFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = ZZDesPago
                ZEntrega = ZZDirentrega
                ZObservaciones1 = Left$(ZZObservaciones, 50)
                ZObservaciones2 = Right$(ZZObservaciones, 50)
                ZOrden = ZZOrden
                ZArticulo = ""
                ZDescripcion = ""
                ZPrecio = "0"
                ZCantidad = "0"
                ZEnvase = ImpreEnvase(Ciclo)
                
                spImprePed = "INSERT INTO ImprePed (" + _
                        "Clave ," + _
                        "Tipo , Pedido ," + _
                        "Renglon , Empresa ," + _
                        "Version , Cliente ," + _
                        "Nombre , Fecha ," + _
                        "Fechaent , TipoPedido ," + _
                        "Condicion , Entrega ," + _
                        "Observaciones1 , Observaciones2 ," + _
                        "Orden , Articulo ," + _
                        "Descripcion , Precio ," + _
                        "Cantidad , Envase )" + _
                        "Values (" + _
                        "'" + ZClave + "'," + _
                        "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                        "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                        "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                        "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                        "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                        "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                        "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                        "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                        "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
                Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
                
        End If
            
    Next A
    
    SumaEspe = 0
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
        SumaEspe = SumaEspe + 1
        
        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = Pedido.Text
        ZRenglon = Str$(WRenglon)
        ZEmpresa = WNombreEmpresa
        ZVersion = ZZVersion
        ZCliente = Cliente.Text
        ZNombre = DesCliente.Caption
        ZFecha = Fecha.Text
        ZFechaent = ZZFecEntrega
        ZTipoPedido = WTipoPedido
        ZCondicion = ZZDesPago
        ZEntrega = ZZDirentrega
        ZObservaciones1 = Left$(ZZObservaciones, 50)
        ZObservaciones2 = Right$(ZZObservaciones, 50)
        ZOrden = ZZOrden
        ZArticulo = ""
        ZDescripcion = WEspecif(SumaEspe)
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
                        
        spImprePed = "INSERT INTO ImprePed (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePed SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "'"
    spImprePed = ZSql
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    If ZZTipoped = 5 Or ZZTipoped = 6 Then
        Listado.ReportFileName = "ImprepedsqlMuestra.rpt"
            Else
        Listado.ReportFileName = "Imprepedsqlsp.rpt"
    End If
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub



Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
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
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
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
        
            Indice = Pantalla.ListIndex
            WEnvases = WIndice.List(Indice)
            spEnvases = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
            
                Entra = "N"
                
                If Val(Envase1.Text) = 0 And Entra = "N" Then
                    Envase1.Text = rstEnvases!Envases
                    Descri1.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti1.SetFocus
                End If
                
                If Val(Envase2.Text) = 0 And Entra = "N" Then
                    Envase2.Text = rstEnvases!Envases
                    Descri2.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti2.SetFocus
                End If
                
                If Val(Envase3.Text) = 0 And Entra = "N" Then
                    Envase3.Text = rstEnvases!Envases
                    Descri3.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti3.SetFocus
                End If
                
                If Val(Envase4.Text) = 0 And Entra = "N" Then
                    Envase4.Text = rstEnvases!Envases
                    Descri4.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti4.SetFocus
                End If
                
                If Val(Envase5.Text) = 0 And Entra = "N" Then
                    Envase5.Text = rstEnvases!Envases
                    Descri5.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti5.SetFocus
                End If
                
                rstEnvases.Close
                    
            End If
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
    
End Sub


Rem
Rem Controles de la wvector1
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
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
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
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
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
            WTexto3.Visible = True
            WTexto3.SetFocus
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
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
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
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
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
            Call StartEdit

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
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
                Call StartEdit
            End If
    
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
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
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
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
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
            Call StartEdit

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
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
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
    StartEdit
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

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 11
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
        Case 3, 4, 5
            
            WVector1.TextMatrix(WVector1.Row, 5) = Pusing("###,###.##", Str$(Val(WVector1.Text)))
            WVector1.Col = 1
            XTerminado = WVector1.Text
            WVector1.Col = 3
            XOriginal = Val(WVector1.Text)
            WVector1.Col = 4
            XCantidad = Val(WVector1.Text)
            WRow = WVector1.Row
            
            ZDife = XCantidad - XOriginal
            ZMargen = XOriginal * 0.25
            
            ZZPasaFila = WVector1.Row
            ZZPasaColumna = WVector1.Col
            ZZPasaProcesoActualiza = "S"
            
            Erase ZZTrabajaLote
            For ZZCiclo = 1 To 100
                ZZTrabajaLote(ZZCiclo) = XLote(WVector1.Row, ZZCiclo)
            Next ZZCiclo
            
            WPasaTerminado = WVector1.TextMatrix(WVector1.Row, 1)
            WPasaCantidad = Val(WVector1.TextMatrix(WVector1.Row, 4))
            WPasaTipoPedido = WTipoPedido
            WPasaCliente = Cliente.Text
            
            PrgMuestraPedidoII.Show
            WControl = "N"
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
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

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cant. S/Pedido"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Cant. Entregar"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Marca"
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

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
































