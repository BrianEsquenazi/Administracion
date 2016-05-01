VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PrgOrdenImpoAnterior 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Ordenes de Compra"
   ClientHeight    =   8340
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   11760
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11760
   Visible         =   0   'False
   Begin VB.Frame DatosImpo 
      Height          =   3975
      Left            =   7440
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Flete 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   55
         Top             =   3360
         Width           =   2175
      End
      Begin VB.ComboBox TipoPago 
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
         Left            =   1440
         TabIndex        =   53
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox PedidoImpo 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   45
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Origen 
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   44
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox Leyenda 
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
         Left            =   1440
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox TipoImpo 
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
         Left            =   1440
         TabIndex        =   42
         Top             =   2280
         Width           =   2175
      End
      Begin MSMask.MaskEdBox FechaImpo 
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Top             =   1800
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
      Begin VB.Label Label31 
         Caption         =   "Flete U$S"
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
         Left            =   240
         TabIndex        =   56
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "Tipo Pago"
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
         Left            =   240
         TabIndex        =   54
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "Condicion"
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
         Left            =   240
         TabIndex        =   51
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Via"
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
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label27 
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
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label26 
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
         Left            =   240
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Origen"
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
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton ImpreRed 
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
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Carpeta 
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
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   39
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox TipoOrden 
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
      Left            =   9840
      TabIndex        =   37
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Moneda 
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
      Left            =   6720
      TabIndex        =   35
      Top             =   120
      Width           =   1935
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
      Left            =   5040
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   6615
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
      Left            =   1080
      MaxLength       =   11
      TabIndex        =   21
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "orden.rpt"
      Destination     =   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   17
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
      TabIndex        =   19
      Top             =   6960
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
      Left            =   5760
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   15
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
   Begin VB.TextBox Orden 
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
      TabIndex        =   13
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
      Top             =   6360
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
      Top             =   6960
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
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   11655
      Begin MSMask.MaskEdBox WFecha2 
         Height          =   300
         Left            =   8160
         TabIndex        =   23
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox WFecha1 
         Height          =   300
         Left            =   6960
         TabIndex        =   22
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
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
         Height          =   300
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   600
         Width           =   1095
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
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion"
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
         Left            =   9360
         TabIndex        =   32
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult. Fecha"
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
         Left            =   8160
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1ra Fecha"
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
         Left            =   6960
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label WPrecio 
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
         Height          =   300
         Left            =   6000
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.Label WCondicion 
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
         Left            =   9360
         TabIndex        =   24
         Top             =   600
         Width           =   1935
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
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   3495
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
      Top             =   6960
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "ordenimpoanterior.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   11655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9480
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
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
      ItemData        =   "ordenimpoanterior.frx":09E6
      Left            =   5040
      List            =   "ordenimpoanterior.frx":09ED
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   6615
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
      TabIndex        =   0
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton LeePedido 
      Caption         =   "Lee     Pedidos"
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
      TabIndex        =   52
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label24 
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
      Left            =   5400
      TabIndex        =   38
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Tipo Orden"
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
      Left            =   8760
      TabIndex        =   36
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Moneda"
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
      TabIndex        =   34
      Top             =   120
      Width           =   1815
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
      Left            =   2640
      TabIndex        =   17
      Top             =   480
      Width           =   2655
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
      TabIndex        =   16
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
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgOrdenImpoAnterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 7 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Precio As Double
Private Condicion As String
Private XMoneda As Integer
Private WMonedaOrden As String
Private WTipoOrden As String
Private WTipoPago As String
Private Cantidad As Single
Private WOrdenprecio As String
Private XPrecio As String
Private XCantidad As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim XParam As String
Dim WEmail As String
Dim Vector(100, 3) As String
Dim ZVector(100, 5) As String
Dim XPorceDerechos(100) As String
Private TipoConsulta As String
Private XVector(10, 5) As String
Private Auxi As String
Private WAuxi As String
Private XEmpresa As String
Private WSaldo As Double
Private Desdelugar As Integer
Dim Tabla(10000) As String
Private WEntre As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim Paridad As Double
Dim WDerechos As String
Dim XDerechos As Double
Dim ZDespacho As Double

Dim ZZFechaLlegada As String
Dim ZZPagoDespacho As String
Dim ZZImpoDespacho As String
Dim ZZVtoDespacho As String
Dim ZZPagoLetra As String
Dim ZZImpoLetra As String
Dim ZZVtoLetra As String


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
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus
    
End Sub



Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgOrdenImpo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Consulta_Click()

    TipoConsulta = "0"

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"

     Opcion.Visible = True
     
 End Sub

Private Sub ImpreRed_Click()

    Renglon = 0
        
    DBGrid1.Refresh
        
    For A = 0 To 9
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
                    
            If Articulo <> "" Then
                        
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                    
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = Articulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                    End If
                        
                    XParam = "'" + Articulo + "','" _
                            + WDescripcion + "'"
                        
                    spArticulo = "ModificaArticuloDescriComercial " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                      
            End If
                                        
        Next iRow
            
    Next A

    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "OrdenImpreImpo.rpt"
    Orden.SetFocus
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT " + _
                            "Orden.Clave, Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Fecha1, Orden.Condicion, " + _
                            "Articulo.Descripcion, Proveedor.Nombre " + _
                        "From " + _
                            DSQ + ".dbo.Orden Orden, " + _
                            DSQ + ".dbo.Articulo Articulo, " + _
                            DSQ + ".dbo.Proveedor Proveedor " + _
                        "Where " + _
                            "Orden.Articulo = Articulo.Codigo AND " + _
                            "Orden.Proveedor = Proveedor.Proveedor AND " + _
                            "Orden.Orden >= " + Orden.Text + " AND " + _
                            "Orden.Orden <= " + Orden.Text + " "
                            
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.Action = 1
    AAAAA = 1

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Liscot
    OPEN_FILE_ImpCtaCtePrv
    
    Select Case WProcesoOrden
        Case 1
            For A = 0 To 9
                Suma = A * 10
                DBGrid1.FirstRow = Suma
                For iRow = 0 To 9
                    For iCol = 0 To 6
                        DBGrid1.Col = iCol
                        DBGrid1.Row = iRow
                        DBGrid1.Text = ""
                    Next iCol
                Next iRow
            Next A
    
            Renglon = 0
            Erase Vector
            
            Rem XEmpresa = WEmpresa
        
            Rem WEmpresa = "0001"
            Rem txtOdbc = "Empresa01"
            Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            Rem For Ciclo = 1 To 100
            Rem     If WVectorOrden(Ciclo, 1) <> "" Then
            Rem
            Rem         ZArticulo = WVectorOrden(Ciclo, 1)
            Rem         WPrecio = 0
            Rem         WCondicion = ""
            Rem
            Rem         spCotiza = "ListaCotizaProveedor " + "'" + WProveedorOrden + "'"
            Rem         Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
            Rem         If rstCotiza.RecordCount > 0 Then
            Rem             With rstCotiza
            Rem                 .MoveFirst
            Rem                 Do
            Rem                     If .EOF = False Then
            Rem
            Rem                         If ZArticulo = rstCotiza!Articulo Then
            Rem                             If rstCotiza!FechaOrd > WFecha Then
            Rem                                 WPrecio = rstCotiza!Precio
            Rem                                 WCondicion = rstCotiza!Condicion
            Rem                             End If
            Rem                         End If
            Rem
            Rem                         .MoveNext
            Rem                             Else
            Rem                         Exit Do
            Rem                     End If
            Rem                 Loop
            Rem             End With
            Rem             rstCotiza.Close
            Rem         End If
                    
            Rem         WVectorOrden(Ciclo, 4) = Str$(WPrecio)
            Rem         WVectorOrden(Ciclo, 5) = WCondicion
            Rem
            Rem     End If
            Rem Next Ciclo
            
            Rem Call Conecta_Empresa
                
            ZGraba = "N"
            
            For Ciclo = 1 To 100
                If WVectorOrden(Ciclo, 1) <> "" Then
            
                    ZGraba = "S"
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = WVectorOrden(Ciclo, 1)
                    Auxi1 = WVectorOrden(Ciclo, 1)
                    
                    spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", WVectorOrden(Ciclo, 2))
                
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", WVectorOrden(Ciclo, 4))
                
                    DBGrid1.Col = 4
                    DBGrid1.Text = Fecha.Text
                
                    DBGrid1.Col = 5
                    DBGrid1.Text = Fecha.Text
                
                    Leyenda.ListIndex = Val(WVectorOrden(Ciclo, 7))
                    DBGrid1.Col = 6
                    DBGrid1.Text = Leyenda.Text
                
                    Vector(Renglon, 1) = Auxi1
                    
                    Origen.Text = WVectorOrden(Ciclo, 6)
                    Leyenda.ListIndex = Val(WVectorOrden(Ciclo, 7))
                    PedidoImpo.Text = WVectorOrden(Ciclo, 8)
                    FechaImpo.Text = WVectorOrden(Ciclo, 9)
                    TipoImpo.ListIndex = Val(WVectorOrden(Ciclo, 10))
                    
                End If
            
            Next Ciclo
    
            WRenglon = Renglon
            Renglon = 0
    
            DBGrid1.FirstRow = 0
    
            Renglon = Renglon + 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
    
            Renglon = Renglon - 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            
            If ZGraba = "S" Then
                Call Graba_Click
            End If
    
        Case Else
    End Select
    WProcesoOrden = 0
End Sub

Private Sub LeePedido_Click()
        WProveedorOrden = Proveedor.Text
    WDesProveedorOrden = DesProveedor.Caption
    PrgOrdenII.Show
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            XEmpresa = WEmpresa
        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spProveedor = "ListaProveedoresOrdConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = RstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = RstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
            
            Call Conecta_Empresa
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            spArticulo = "ListaArticuloConsulta"
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
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 10 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -   -   "
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = DBGrid1.Text
            Else
        WCantidad.Text = ""
    End If
    
    DBGrid1.Col = 3
    WPrecio.Caption = DBGrid1.Text
    
    DBGrid1.Col = 4
    If DBGrid1.Text <> "" Then
        WFecha1.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 5
    If DBGrid1.Text <> "" Then
        WFecha2.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 6
    WCondicion.Caption = DBGrid1.Text
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha de la orden de compra es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
        If TipoOrden.ListIndex = 1 Then
        If TipoImpo.ListIndex = 0 Then
            m$ = "Se debe informar la via de transporte"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
    End If
    
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 0 Then
            m$ = "Se debe informar la condicion de la importacion"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Exit Sub
        End If
    End If
            
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 1 Then
            If Val(Flete.Text) = 0 Then
                m$ = "Se debe informar el monto del flete"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                Exit Sub
            End If
        End If
    End If
    
    ZParidad = 0
    If TipoOrden.ListIndex = 1 Then
    
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
    
        XXFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        spCambios = "ConsultaCambio  " + "'" + XXFecha + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZParidad = rstCambios!Cambio
            rstCambios.Close
            Call Conecta_Empresa
                    Else
            m$ = "Se debe informar la paridad"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Call Conecta_Empresa
            Exit Sub
        End If
        
    End If

    
    XEmpresa = WEmpresa
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    ZZFechaLlegada = "  /  /    "
    ZZPagoDespacho = "0"
    ZZImpoDespacho = "0"
    ZZVtoDespacho = "  /  /    "
    ZZPagoLetra = "0"
    ZZImpoLetra = "0"
    ZZVtoLetra = "  /  /    "
    ZZAuxiFecha = "  /  /    "

    Rem Borra la Ordenes anteriores
    
    Renglon = 0
    Erase Vector
    
    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstOrden.RecordCount > 0 Then
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
            
                Vector(Renglon, 1) = rstOrden!Articulo
                Vector(Renglon, 2) = Str$(rstOrden!Cantidad)
                XDerechos = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                Vector(Renglon, 3) = Str$(XDerechos)
                
                ZZFechaLlegada = IIf(IsNull(rstOrden!FechaLlegada), "  /  /    ", rstOrden!FechaLlegada)
                ZZPagoDespacho = IIf(IsNull(rstOrden!PagoDespacho), "0", rstOrden!PagoDespacho)
                ZZImpoDespacho = IIf(IsNull(rstOrden!ImpoDespacho), "0", rstOrden!ImpoDespacho)
                ZZVtoDespacho = IIf(IsNull(rstOrden!VtoDespacho), "  /  /    ", rstOrden!VtoDespacho)
                ZZPagoLetra = IIf(IsNull(rstOrden!PagoLetra), "0", rstOrden!PagoLetra)
                ZZImpoLetra = IIf(IsNull(rstOrden!ImpoLetra), "0", rstOrden!ImpoLetra)
                ZZVtoLetra = IIf(IsNull(rstOrden!VtoLetra), "  /  /    ", rstOrden!VtoLetra)

                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    End If
    
    For Da = 1 To Renglon
    
        Articulo = Vector(Da, 1)
        Cantidad = Vector(Da, 2)
    
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
                
            WCodigo = Articulo
            WCosto1 = Str$(rstArticulo!Costo1)
            WFecha = ""
            WFecha = rstArticulo!Fecha
            WOrden = Str$(rstArticulo!Orden)
            WPedido = Str$(rstArticulo!Pedido - Val(Cantidad))
            WProveedor = ""
            WProveedor = rstArticulo!Proveedor
            WDate = Date$
            
            rstArticulo.Close
                        
            XParam = "'" + WCodigo + "','" _
                    + WCosto1 + "','" _
                    + WPedido + "','" _
                    + WFecha + "','" _
                    + WOrden + "','" _
                    + WProveedor + "','" _
                    + WDate + "'"
                        
            spArticulo = "ModificaArticuloOrden " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        WCantot = Val(Cantidad)
        WMarca = ""
        WLugar = 0
        Erase Tabla
                
        XParam = "'" + Articulo + "','" _
                        + WMarca + "'"
        spSolic = "ListaSolicitudBajaArticulo " + XParam
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount > 0 Then
                
            With rstSolic
    
                .MoveFirst
                If .NoMatch = False Then
                    Do
                            
                        WLugar = WLugar + 1
                        Tabla(WLugar) = rstSolic!Clave
                            
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
        
            End With
            rstSolic.Close
                    
        End If
                
        For Cicla = WLugar To 1 Step -1
                
            WClave = Tabla(Cicla)
                
            spSolic = "ConsultaSolicitud " + "'" + WClave + "'"
            Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            If rstSolic.RecordCount > 0 Then
            
                WEntregado = rstSolic!Entregado
                rstSolic.Close
                
                If WEntregado <> 0 Then
                        
                    If WEntregado >= WCantot Then
                        WEntregado = WEntregado - WCantot
                        WMarca = ""
                        Salida = "S"
                            Else
                        WCantot = WCantot - WEntregado
                        WEntregado = 0
                        WMarca = ""
                        Salida = "N"
                    End If
                        
                    WEntre = WEntregado
                        
                    XParam = "'" + WClave + "','" _
                            + WEntre + "','" _
                            + WMarca + "'"
                        
                    spSolic = "ModificaSolicitudEntregado " + XParam
                    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                        
                If Salida = "S" Then
                    Exit For
                End If
                
            End If
        Next Cicla
        
    Next Da
        
    spOrden = "BorrarOrdenTotal " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenDynaset, dbSQLPassThrough)
        
    Renglon = 0
    ZSuma = 0
        
    DBGrid1.Refresh
        
    For A = 0 To 9
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = UCase(DBGrid1.Text)
            
            DBGrid1.Col = 2
            Cantidad = Val(DBGrid1.Text)
            XCantidad = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            XPrecio = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Fecha1 = DBGrid1.Text
                    
            DBGrid1.Col = 5
            fecha2 = DBGrid1.Text
            If fecha2 <> "" Then
                ZZAuxiFecha = fecha2
            End If
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
            WWPorceDerechos = XPorceDerechos(WLugar)
                    
            DBGrid1.Col = 6
            Condicion = DBGrid1.Text
                    
            If Articulo <> "" Then
            
                Renglon = Renglon + 1
                ZSuma = ZSuma + (Val(XCantidad) * Val(XPrecio))
                
                ZVector(Renglon, 1) = Articulo
                ZVector(Renglon, 2) = XCantidad
                ZVector(Renglon, 3) = XPrecio
                ZVector(Renglon, 4) = WWPorceDerechos
            
                WOrden = Orden.Text
                WRenglon = Str$(Renglon)
                WFecha = Fecha.Text
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WProveedor = Proveedor.Text
                WArticulo = Articulo
                WCantidad = XCantidad
                WPrecio = XPrecio
                WFecha1 = Fecha1
                WFecha2 = fecha2
                WCondicion = Condicion
                WRecibida = "0"
                XSaldo = "0"
                WLiberada = "0"
                WDevuelta = "0"
                WFechaEntrega = "  /  /    "
                Auxi1 = WOrden
                Auxi = WRenglon
                Call Ceros(Auxi1, 6)
                Call Ceros(Auxi, 2)
                WClave = Auxi1 + Auxi
                WDate = Date$
                WMonedaOrden = Str$(Moneda.ListIndex)
                WTipoOrden = Str$(TipoOrden.ListIndex)
                WTipoPago = Str$(TipoPago.ListIndex)
                WCarpeta = Carpeta.Text
                WDerechos = "0"
                WOrigen = Origen.Text
                
                For Cicla = 1 To 100
                    If WArticulo = Vector(Cicla, 1) Then
                        WDerechos = Vector(Cicla, 3)
                        Exit For
                    End If
                Next Cicla
                         
                XParam = "'" + WClave + "','" _
                         + WOrden + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WProveedor + "','" _
                         + WArticulo + "','" _
                         + WCantidad + "','" _
                         + WPrecio + "','" _
                         + WFecha1 + "','" _
                         + WFecha2 + "','" _
                         + WCondicion + "','" _
                         + WRecibida + "','" _
                         + XSaldo + "','" _
                         + WFechaord + "','" _
                         + WLiberada + "','" _
                         + WDevuelta + "','" _
                         + WFechaEntrega + "','" _
                         + WDate + "','" _
                         + WMonedaOrden + "','" _
                         + WTipoOrden + "','" _
                         + WCarpeta + "','" _
                         + WDerechos + "','" _
                         + WOrigen + "'"
                         
                spOrden = "AltaOrdenIII " + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                
                WLeyenda = Str$(Leyenda.ListIndex)
                XParam = "'" + WOrden + "','" _
                             + WLeyenda + "'"
                spOrden = "ModificaOrdenLeyenda " + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                
                    WCodigo = WArticulo
                    Rem WCosto1 = Str$(Precio)
                    WCosto1 = Str$(rstArticulo!Costo1)
                    WFecha = Fecha.Text
                    WOrden = Orden.Text
                    WPedido = Str$(rstArticulo!Pedido + Cantidad)
                    WProveedor = ""
                    WProveedor = rstArticulo!Proveedor
                    WDate = Date$
                    rstArticulo.Close
                    
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = WArticulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                    End If
                    
                    If TipoOrden.ListIndex = 2 Then
                        WPedido = Str$(Cantidad)
                        XParam = "'" + WCodigo + "','" _
                            + WPedido + "'"
                        
                        spArticulo = "ModificaArticuloPedido " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                    
                        XParam = "'" + WCodigo + "','" _
                            + WCosto1 + "','" _
                            + WPedido + "','" _
                            + WFecha + "','" _
                            + WOrden + "','" _
                            + WProveedor + "','" _
                            + WDate + "'"
                        
                        spArticulo = "ModificaArticuloOrden " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                        XParam = "'" + WCodigo + "','" _
                                + WDescripcion + "'"
                        
                        spArticulo = "ModificaArticuloDescriComercial " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
            
                Rem Actualiza la solicitud de orden de compra

                WCantot = Val(XCantidad)
                WMarca = "X"
                WLugar = 0
                Erase Tabla
                
                XParam = "'" + WArticulo + "','" _
                             + WMarca + "'"
                spSolic = "ListaSolicitudPendienteArticulo " + XParam
                Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                If rstSolic.RecordCount > 0 Then
                
                    With rstSolic
    
                        .MoveFirst
                        If .NoMatch = False Then
                            Do
                            
                                WLugar = WLugar + 1
                                Tabla(WLugar) = rstSolic!Clave
                            
                                .MoveNext
                
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                            Loop
                        End If
        
                    End With
                    rstSolic.Close
                    
                End If
                
                For Cicla = 1 To WLugar
                
                    WClave = Tabla(Cicla)
                
                    spSolic = "ConsultaSolicitud " + "'" + WClave + "'"
                    Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                    If rstSolic.RecordCount > 0 Then
                    
                        WCanti = rstSolic!Cantidad - rstSolic!Entregado
                        WEntregado = rstSolic!Entregado
                        rstSolic.Close
                        

                        If WCanti > WCantot Then
                            WEntregado = WEntregado + WCantot
                            WMarca = ""
                            Salida = "S"
                                Else
                            WEntregado = WEntregado + WCanti
                            WCantot = WCantot - WCanti
                            WMarca = "X"
                            Salida = "N"
                        End If
                        
                        WEntre = WEntregado
                        
                        XParam = "'" + WClave + "','" _
                                + WEntre + "','" _
                                + WMarca + "'"
                        
                        spSolic = "ModificaSolicitudEntregado " + XParam
                        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                        
                        If Salida = "S" Then
                            Exit For
                        End If
                    End If
                Next Cicla
                      
            End If
                                        
        Next iRow
            
    Next A
    
    ZOrigen = Origen.Text
    ZLeyenda = Str$(Leyenda.ListIndex)
    ZPedidoImpo = PedidoImpo.Text
    ZFechaImpo = FechaImpo.Text
    ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
    ZTipoImpo = Str$(TipoImpo.ListIndex)
    ZTipoPago = Str$(TipoPago.ListIndex)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " Origen = " + "'" + ZOrigen + "',"
    ZSql = ZSql + " Leyenda = " + "'" + ZLeyenda + "',"
    ZSql = ZSql + " PedidoImpo = " + "'" + ZPedidoImpo + "',"
    ZSql = ZSql + " FechaImpo = " + "'" + ZFechaImpo + "',"
    ZSql = ZSql + " OrdFechaImpo = " + "'" + ZOrdFechaImpo + "',"
    ZSql = ZSql + " TipoImpo = " + "'" + ZTipoImpo + "',"
    ZSql = ZSql + " TipoPago = " + "'" + ZTipoPago + "'"
    
    ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    WOrden = Orden.Text
    WImpresion = "N"
    XParam = "'" + WOrden + "','" _
                 + WImpresion + "'"
    spOrden = "ModificaOrdenImpresion " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                
    T$ = "Orden de Compra"
    m$ = "Desea imprimir la Orden de Compra"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call ImpreRed_Click
    End If
        
    Call Conecta_Empresa
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Orden.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""
    
    WArticulo.SetFocus
    
End Sub

Private Sub Limpia_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""

    Orden.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    WEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    
    Moneda.ListIndex = 0
    TipoOrden.ListIndex = 1
    TipoPago.ListIndex = 0
    Leyenda.ListIndex = 0
    TipoImpo.ListIndex = 0
    
    For A = 0 To 9
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Rem With rstOrden
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Orden.Text = !Orden + 1
    Rem             Else
    Rem         Orden.Text = ""
    Rem     End If
    Rem End With
    
    Orden.Text = "1"
    
    Rem spOrden = "ListaOrdenNUmero"
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstOrden.RecordCount > 0 Then
    Rem     With rstOrden
    Rem         .MoveLast
    Rem         Orden.Text = rstOrden!Orden + 1
    Rem     End With
    Rem     rstOrden.Close
    Rem         Else
    Rem     Orden.Text = "1"
    Rem End If
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden < 700000"
    ZSql = ZSql + " Order by Orden.Clave"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveLast
            Orden.Text = Str$(rstOrden!Orden + 1)
        End With
        rstOrden.Close
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    Graba.Enabled = True

    Orden.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ingre = "N"
        WArticulo.Text = UCase(WArticulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            Ingre = "S"
            rstArticulo.Close
                Else
            WArticulo.SetFocus
        End If
        If Ingre = "S" Then
            If TipoOrden.ListIndex <> 2 Then
                Call Calcula_Precio(Proveedor.Text, WArticulo.Text, Precio, Condicion, XMoneda)
                If Precio = 0 Then
                    Desdelugar = 1
                    XCoti.Height = 4200
                    XCoti.Left = 2360
                    XCoti.Top = 1320
                    XCoti.Width = 7455
                    XCoti.Visible = True
                    XProve.Text = Proveedor.Text
                    XDesProve.Caption = DesProveedor.Caption
                    XArti = WArticulo.Text
                    XPrec.Text = ""
                    XCondicion.Text = ""
                    XObservaciones.Text = ""
                    XPrec.SetFocus
                        Else
                    If Moneda.ListIndex = 2 Then
                        Moneda.ListIndex = XMoneda
                    End If
                    If Moneda.ListIndex = XMoneda Then
                        WPrecio.Caption = Pusing("###,###.##", Str$(Precio))
                        WCondicion.Caption = Condicion
                        WCantidad.SetFocus
                            Else
                        m$ = "La moneda de la cotizacion no se corresponde a la moneda de los otros productos cargados en esta orden de compra"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        WArticulo.SetFocus
                    End If
                End If
                    Else
                WPrecio.Caption = ""
                WPrecio.Caption = Pusing("###,###.##", Str$(Precio))
                WCondicion.Caption = ""
                WCantidad.SetFocus
            End If
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WFecha1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WFecha1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(WFecha1.Text, Auxi)
        If Auxi = "S" Then
            WFecha2.SetFocus
                Else
            WFecha1.SetFocus
        End If
    End If
End Sub

Private Sub WFecha2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(WFecha2.Text, Auxi)
        If Auxi = "S" Then
        
            Call Valida_fecha(WFecha1.Text, Auxi)
            If Auxi <> "S" Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha1.SetFocus
                Exit Sub
            End If
        
            Call Valida_fecha(WFecha2.Text, Auxi)
            If Auxi <> "S" Then
                m$ = "La fecha de entrega prevista es incorrecta"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                WFecha2.SetFocus
                Exit Sub
            End If
        
            Rem Call Alta_Vector
            Rem Call Ingresa_Click
            WArticulo.SetFocus
            
                Else
                
            WFecha2.SetFocus
            
        End If
    End If
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Select Case Val(TipoConsulta)
                    Case 2
                        XProv1.Text = WProveedor
                        XDesProv1.Caption = RstProveedor!Nombre
                        XProv1.SetFocus
                    Case 3
                        XProve.Text = WProveedor
                        XDesProve.Caption = RstProveedor!Nombre
                        XProve.SetFocus
                    Case 4
                        XProv3.Text = WProveedor
                        XDesProv3.Caption = RstProveedor!Nombre
                        XProv3.SetFocus
                    Case Else
                        Proveedor.Text = WProveedor
                        DesProveedor.Caption = RstProveedor!Nombre
                        WEmail = RstProveedor!email
                        Proveedor.SetFocus
                End Select
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
        
                Select Case Val(TipoConsulta)
                    Case 2
                        XArt2.Text = rstArticulo!Codigo
                        XDesArt2.Caption = rstArticulo!Descripcion
                        XArt2.SetFocus
                        
                    Case 3
                        XArti.Text = rstArticulo!Codigo
                        XDesArti.Caption = rstArticulo!Descripcion
                        XArti.SetFocus
                    
                    Case Else
                        WArticulo.Text = rstArticulo!Codigo
                        WDescripcion.Caption = rstArticulo!Descripcion
                    
                        DBGrid1.Col = 0
                        DBGrid1.Text = rstArticulo!Codigo
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstArticulo!Descripcion
                    
                        Call Alta_Vector
                        WLinea.Text = WAnterior + 1
                        If Val(WLinea.Text) > 0 Then
                            DBGrid1.Row = Val(WLinea.Text) - 1
                        End If
                    
                        Call DBGrid1.SetFocus
                        WCantidad.SetFocus
                        
                End Select
                    
            End If
            
        Case Else
    End Select
    
    Ayuda.Visible = False
    Pantalla.Visible = False
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 100 Then
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
ReDim UserData(0 To 6, 0 To 100)

mTotalRows& = 100

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
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3500
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 4
             DBGrid1.Columns(newcnt).Caption = "1ra Fecha"
             DBGrid1.Columns(newcnt).Width = 1150
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Ult Fecha"
             DBGrid1.Columns(newcnt).Width = 1150
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Condicion"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    TipoImpo.Clear
    
    TipoImpo.AddItem ""
    TipoImpo.AddItem "Maritimo"
    TipoImpo.AddItem "Terrestre"
    TipoImpo.AddItem "Areo"
    
    TipoImpo.ListIndex = 0
 
    WLinea.Text = ""
    WArticulo.Text = "  -   -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WFecha1.Text = "  /  /    "
    WFecha2.Text = "  /  /    "
    WCondicion.Caption = ""

    Orden.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    WEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    TipoImpo.ListIndex = 0
    
    Leyenda.Clear
    
    Leyenda.AddItem ""
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CIF"
    Leyenda.AddItem "CFR"
    Leyenda.AddItem "CPT"
    Leyenda.AddItem "EXW"
    Leyenda.AddItem "FCA"
    
    Leyenda.ListIndex = 0
 
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"
    Moneda.AddItem ""

    Moneda.ListIndex = 0
    
    TipoOrden.Clear
    
    TipoOrden.AddItem "Local"
    TipoOrden.AddItem "Importacion"
    TipoOrden.AddItem "Prestamo"

    TipoOrden.ListIndex = 1
    
    TipoPago.Clear
    
    TipoPago.AddItem ""
    TipoPago.AddItem "Pago Anticipado"
    TipoPago.AddItem "A la vista"
    TipoPago.AddItem "Cuenta Corriente"

    TipoPago.ListIndex = 0
 
    Orden.Text = "1"
    
    Rem spOrden = "ListaOrdenNUmero"
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstOrden.RecordCount > 0 Then
    Rem     With rstOrden
    Rem         .MoveLast
    Rem         Orden.Text = rstOrden!Orden + 1
    Rem     End With
    Rem     rstOrden.Close
    Rem         Else
    Rem     Orden.Text = "1"
    Rem End If
    
    ZSql = ""
    ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden < 700000"
    ZSql = ZSql + " Order by Orden.Clave"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveLast
            Orden.Text = Str$(rstOrden!Orden + 1)
        End With
        rstOrden.Close
    End If

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdenImpo.Caption = "Ingreso de Orden de Compras :  " + !Nombre
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Graba.Enabled = True
    
    Orden.SetFocus
    
End Sub

Private Sub Proceso_Click()

    Graba.Enabled = True
    
    For A = 0 To 9
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Vector
    
    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstOrden
        .MoveFirst
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                
                DBGrid1.Col = 0
                DBGrid1.Text = rstOrden!Articulo
                Auxi1 = rstOrden!Articulo
                
                DBGrid1.Col = 2
                DBGrid1.Text = Pusing("###,###.##", rstOrden!Cantidad)
                
                DBGrid1.Col = 3
                DBGrid1.Text = Pusing("###,###.##", rstOrden!Precio)
                
                DBGrid1.Col = 4
                DBGrid1.Text = rstOrden!Fecha1
                
                DBGrid1.Col = 5
                DBGrid1.Text = rstOrden!fecha2
                
                DBGrid1.Col = 6
                DBGrid1.Text = rstOrden!Condicion
                
                If rstOrden!Recibida > 0 Then
                    Graba.Enabled = False
                End If
                
                Vector(Renglon, 1) = Auxi1
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    
    If Graba.Enabled = False Then
        m$ = "La orden de compra no podra ser actualizada ya que posee productos que fueron cumplidos en forma total o parcial"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
            
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        Auxi1 = Vector(Renglon, 1)
    
        spArticulo = "ConsultaArticulo " + "'" + Auxi1 + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstArticulo!Descripcion
            WArticulo.SetFocus
            rstArticulo.Close
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
    
    WArticulo.SetFocus

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
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = WFecha1.Text
            
            DBGrid1.Col = 5
            DBGrid1.Text = WFecha2.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WCondicion.Caption
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
            
            DBGrid1.Col = 3
            DBGrid1.Text = Pusing("###,###.##", WPrecio.Caption)
            
            DBGrid1.Col = 4
            DBGrid1.Text = WFecha1.Text
            
            DBGrid1.Col = 5
            DBGrid1.Text = WFecha2.Text
                        
            DBGrid1.Col = 6
            DBGrid1.Text = WCondicion.Caption
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        WClave = Auxi + "01"
            
        Entra = "N"
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            Origen.Text = IIf(IsNull(rstOrden!Origen), "", rstOrden!Origen)
            Carpeta.Text = IIf(IsNull(rstOrden!Carpeta), "", rstOrden!Carpeta)
            Moneda.ListIndex = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            TipoOrden.ListIndex = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            TipoPago.ListIndex = IIf(IsNull(rstOrden!TipoPago), "0", rstOrden!TipoPago)
            Leyenda.ListIndex = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
            PedidoImpo.Text = IIf(IsNull(rstOrden!PedidoImpo), "", rstOrden!PedidoImpo)
            FechaImpo.Text = IIf(IsNull(rstOrden!FechaImpo), "  /  /    ", rstOrden!FechaImpo)
            TipoImpo.ListIndex = IIf(IsNull(rstOrden!TipoImpo), "0", rstOrden!TipoImpo)
            Fecha.Text = rstOrden!Fecha
            Proveedor.Text = rstOrden!Proveedor
            rstOrden.Close
            Entra = "S"
                Else
            WOrden = Orden.Text
            Call Limpia_Click
            Orden.Text = WOrden
            Fecha.SetFocus
        End If
        
        If Entra = "S" Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                DesProveedor.Caption = RstProveedor!Nombre
                WEmail = RstProveedor!email
                RstProveedor.Close
            End If
            Call Proceso_Click
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                    Proveedor.Text = RstProveedor!Proveedor
                    DesProveedor.Caption = RstProveedor!Nombre
                    WEmail = RstProveedor!email
                    If TipoOrden.ListIndex = 1 Then
                        Carpeta.SetFocus
                            Else
                        WArticulo.SetFocus
                    End If
                        Else
                    Proveedor.SetFocus
            End If
                Else
            TipoConsulta = "1"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
End Sub

Private Sub Carpeta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DatosImpo.Visible = True
        Origen.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Leyenda.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Leyenda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PedidoImpo.SetFocus
    End If
End Sub

Private Sub PedidoImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaImpo.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaImpo.Text, Auxi)
        If Auxi = "S" Then
            TipoImpo.SetFocus
                Else
            FechaImpo.SetFocus
        End If
    End If
End Sub

Private Sub TipoImpo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoPago.SetFocus
    End If
End Sub

Private Sub TipoPago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Flete.SetFocus
    End If
End Sub

Private Sub Flete_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        DatosImpo.Visible = False
        
        ZOrigen = Origen.Text
        ZLeyenda = Str$(Leyenda.ListIndex)
        ZPedidoImpo = PedidoImpo.Text
        ZFechaImpo = FechaImpo.Text
        ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
        ZTipoImpo = Str$(TipoImpo.ListIndex)
        ZTipoPago = Str$(TipoPago.ListIndex)
        ZFlete = Flete.Text
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Orden SET "
        ZSql = ZSql + " Flete = " + "'" + ZFlete + "',"
        ZSql = ZSql + " Origen = " + "'" + ZOrigen + "',"
        ZSql = ZSql + " Leyenda = " + "'" + ZLeyenda + "',"
        ZSql = ZSql + " PedidoImpo = " + "'" + ZPedidoImpo + "',"
        ZSql = ZSql + " FechaImpo = " + "'" + ZFechaImpo + "',"
        ZSql = ZSql + " OrdFechaImpo = " + "'" + ZOrdFechaImpo + "',"
        ZSql = ZSql + " TipoImpo = " + "'" + ZTipoImpo + "',"
        ZSql = ZSql + " TipoPago = " + "'" + ZTipoPago + "'"
        ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
        
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        
        WArticulo.SetFocus
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta_DblClick()
    DatosImpo.Visible = True
    Origen.SetFocus
End Sub


Sub Calcula_Precio(WProveedor, WArticulo As String, WPrecio As Double, WCondicion As String, WMoneda As Integer)

    WPrecio = 0
    WCondicion = ""
    WFecha = ""
    
    XEmpresa = WEmpresa
        
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spCotiza = "ListaCotizaProveedor " + "'" + WProveedor + "'"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstCotiza.RecordCount > 0 Then
    With rstCotiza
        .MoveFirst
        Do
            If .EOF = False Then
            
                If WArticulo = rstCotiza!Articulo Then
                    
                    If rstCotiza!FechaOrd > WFecha Then
                        WPrecio = rstCotiza!Precio
                        WCondicion = rstCotiza!Condicion
                        WCotiza = rstCotiza!Cotiza
                        WFecha = rstCotiza!FechaOrd
                        WMoneda = rstCotiza!Moneda
                    End If
                End If

                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstCotiza.Close
    End If
    
    Call Conecta_Empresa
    
    A = 1

End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
        
            XEmpresa = WEmpresa
        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            If RstProveedor.RecordCount > 0 Then
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        Da = Len(RstProveedor!Nombre) - WEspacios
                
                        For aa = 1 To Da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                Auxi = Str$(RstProveedor!Proveedor)
                                Call Ceros(Auxi, 11)
                                IngresaItem = Auxi + "    " + RstProveedor!Nombre
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Proveedor
                                WIndice.AddItem IngresaItem
                                Exit For
                            End If
                        Next aa
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
            End If
            
            Call Conecta_Empresa
            
            
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
    
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            Da = Len(rstArticulo!Descripcion) - WEspacios
                
                            For Aaa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
    
                rstArticulo.Close
            End If
                
            
        Case Else
        
    End Select
    
    End If

End Sub

Private Sub XAcepta_Click()

    XEmpresa = WEmpresa
        
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WCotiza = "1"
    
    spCotiza = "ListaCotizaNumero"
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
        With rstCotiza
            .MoveLast
            WCotiza = rstCotiza!Cotiza + 1
        End With
        rstCotiza.Close
    End If

    Articulo = XArti.Text
    Precio = XPrec.Text
    Condicion = XCondicion.Text
    Observaciones = XObservaciones.Text
    XRenglon = 1
    
    Auxi = Str$(XRenglon)
    Call Ceros(Auxi, 2)
    Auxi1 = Str$(WCotiza)
    Call Ceros(Auxi1, 6)
                        
    WCot = Str$(WCotiza)
    WRenglon = Str$(XRenglon)
    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WProveedor = XProve.Text
    WArticulo = XArti.Text
    WPrecio = XPrec.Text
    WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WCondicion = XCondicion.Text
    WObservaciones = XObservaciones.Text
    WClave = Auxi1 + Auxi
    WDate = Date$
    WMoneda = Str$(Moneda3.ListIndex)
        
    XParam = "'" + WClave + "','" _
                + WCot + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WProveedor + "','" _
                + WArticulo + "','" _
                + WPrecio + "','" _
                + WCondicion + "','" _
                + WObservaciones + "','" _
                + WFechaord + "','" _
                + WDate + "','" _
                + WMoneda + "'"
                    
    spCotiza = "AltaCotizaII " + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Conecta_Empresa
    
    XCoti.Visible = False
    If Desdelugar = 0 Then
        Orden.SetFocus
            Else
        Rem WCantidad.Text = ""
        WCantidad.SetFocus
    End If
    
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

