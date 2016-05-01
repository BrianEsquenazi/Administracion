VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenImpo 
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
      TabIndex        =   60
      Top             =   7560
      Width           =   975
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
      TabIndex        =   59
      Top             =   7560
      Width           =   975
   End
   Begin VB.Frame IngreDerechos 
      Height          =   1095
      Left            =   3360
      TabIndex        =   56
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WPorceDerechos 
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
         Left            =   720
         MaxLength       =   10
         TabIndex        =   57
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Derechos"
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
         Left            =   720
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame DatosImpo 
      Height          =   3855
      Left            =   7440
      TabIndex        =   40
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
         TabIndex        =   54
         Top             =   3240
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
         TabIndex        =   52
         Top             =   2760
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   2280
         Width           =   2175
      End
      Begin MSMask.MaskEdBox FechaImpo 
         Height          =   285
         Left            =   1440
         TabIndex        =   45
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
         TabIndex        =   55
         Top             =   3240
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
         TabIndex        =   53
         Top             =   2760
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
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
      Left            =   6360
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
      Top             =   6240
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
      Left            =   5040
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   3855
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
      OleObjectBlob   =   "OrdenImpo.frx":0000
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
      ItemData        =   "OrdenImpo.frx":09E6
      Left            =   5040
      List            =   "OrdenImpo.frx":09ED
      TabIndex        =   1
      Top             =   6600
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
   Begin VB.CommandButton Complemento 
      Caption         =   "Datos Compleme."
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
      TabIndex        =   51
      Top             =   7560
      Width           =   1335
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
Attribute VB_Name = "PrgOrdenImpo"
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
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstMarcas As Recordset
Dim spMarcas As String
Dim XParam As String
Dim ZEmail As String
Dim Vector(100, 3) As String
Dim ZVector(100, 5) As String
Dim XPorceDerechos(100) As String
Private TipoConsulta As String
Private XVector(10, 5) As String
Private Auxi As String
Private WAuxi As String
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

Private Sub Complemento_Click()

    If TipoOrden.ListIndex <> 1 Then
        m$ = "La Orden no es de importacion"
        a% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
        Exit Sub
    End If

    WPasaOrden = Orden.Text
    WPasaCarpeta = Carpeta.Text
    WPasaOrigen = 3
    PrgOrdenComplemento.Show
    Rem Call PrgOrdenComplemento.Orden_Keypress(13)
    
End Sub

Private Sub Consulta_Click()

    TipoConsulta = "0"

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"

     Opcion.Visible = True
     
 End Sub

Private Sub Cotart_Click()

    Moneda2.ListIndex = 0

    XCotart.Height = 1935
    XCotart.Left = 2280
    XCotart.Top = 1800
    XCotart.Width = 8655

    XCotart.Visible = True
    XArt2.Text = "  -   -   "
    XDesArt2.Caption = ""
    XArt2.SetFocus

End Sub

Private Sub Cotprv_Click()

    Moneda1.ListIndex = 0

    XCotPrv.Height = 2300
    XCotPrv.Left = 2280
    XCotPrv.Top = 1800
    XCotPrv.Width = 8655

    XCotPrv.Visible = True
    XProv1.Text = ""
    XDesProv1.Caption = ""
    XProv1.SetFocus

End Sub

Private Sub EMail_Click()
        
    Renglon = 0
        
    DBGrid1.Refresh
        
    For a = 0 To 9
        
        Suma = a * 10
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
            
    Next a

    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 3
    
    Listado.EMailToList = ZEmail
    Listado.EMailSubject = "ORDEN DE COMPRA NUMERO " + Orden.Text
    Listado.EMailMessage = "Se remite por la presente la orden de compra " + Orden.Text
    
    Select Case Val(WEmpresa)
        Case 1
            Listado.ReportFileName = "Orden1.rpt"
        Case 2
            Listado.ReportFileName = "Orden11.rpt"
        Case 3
            Listado.ReportFileName = "Orden2.rpt"
        Case 4
            Listado.ReportFileName = "Orden22.rpt"
        Case 5
            Listado.ReportFileName = "Orden3.rpt"
        Case 6
            Listado.ReportFileName = "Orden4.rpt"
        Case 7
            Listado.ReportFileName = "Orden7.rpt"
        Case 8
            Listado.ReportFileName = "Orden8.rpt"
        Case 9
            Listado.ReportFileName = "Orden9.rpt"
        Case Else
            Listado.ReportFileName = "Orden.rpt"
    End Select

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
    
    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
            
End Sub

Private Sub ImpreRed_Click()

    Renglon = 0
        
    DBGrid1.Refresh
        
    For a = 0 To 9
        
        Suma = a * 10
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
            
    Next a

    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Select Case Val(WEmpresa)
        Case 1
            Listado.ReportFileName = "OrdenImpre1.rpt"
        Case 2
            Listado.ReportFileName = "OrdenImpre11.rpt"
        Case 3
            Listado.ReportFileName = "OrdenImpre2.rpt"
        Case 4
            Listado.ReportFileName = "OrdenImpre22.rpt"
        Case 5
            Listado.ReportFileName = "OrdenImpre3.rpt"
        Case 6
            Listado.ReportFileName = "OrdenImpre4.rpt"
        Case 7
            Listado.ReportFileName = "OrdenImpre7.rpt"
        Case 8
            Listado.ReportFileName = "OrdenImpre8.rpt"
        Case 9
            Listado.ReportFileName = "OrdenImpre9.rpt"
        Case Else
            Listado.ReportFileName = "OrdenImpre.rpt"
    End Select

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


Private Sub CtaCte_Click()

    XCc.Height = 1935
    XCc.Left = 2280
    XCc.Top = 1800
    XCc.Width = 8655

    XCc.Visible = True
    XProv3.Text = ""
    XDesProv3.Caption = ""
    XProv3.SetFocus

End Sub

Private Sub Email2_Click()

    Renglon = 0
        
    DBGrid1.Refresh
        
    For a = 0 To 9
        
        Suma = a * 10
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
            
    Next a

    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
    Listado.WindowTitle = "Emision de Orden de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Orden.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.Destination = 0
    
    Listado.EMailToList = ZEmail
    Listado.EMailSubject = "ORDEN DE COMPRA NUMERO" + Orden.Text
    Listado.EMailMessage = "Se remite por la presente la orden de compra " + Orden.Text
    
    Select Case Val(WEmpresa)
        Case 1
            Listado.ReportFileName = "Orden1.rpt"
        Case 2
            Listado.ReportFileName = "Orden11.rpt"
        Case 3
            Listado.ReportFileName = "Orden2.rpt"
        Case 4
            Listado.ReportFileName = "Orden22.rpt"
        Case 5
            Listado.ReportFileName = "Orden3.rpt"
        Case 6
            Listado.ReportFileName = "Orden4.rpt"
        Case 7
            Listado.ReportFileName = "Orden7.rpt"
        Case 8
            Listado.ReportFileName = "Orden8.rpt"
        Case 9
            Listado.ReportFileName = "Orden9.rpt"
        Case Else
            Listado.ReportFileName = "Orden.rpt"
    End Select
    
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
    
    Rem ChDrive MiRutaII
    Rem ChDir MiRuta
    
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    OPEN_FILE_Empresa
    OPEN_FILE_Liscot
    OPEN_FILE_ImpCtaCtePrv
    
    Select Case WProcesoOrden
        Case 1
            For a = 0 To 9
                Suma = a * 10
                DBGrid1.FirstRow = Suma
                For iRow = 0 To 9
                    For iCol = 0 To 6
                        DBGrid1.Col = iCol
                        DBGrid1.Row = iRow
                        DBGrid1.Text = ""
                    Next iCol
                Next iRow
            Next a
    
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
                    TipoOrden.ListIndex = 1
                    
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


Private Sub Ingrecot_Click()

    Moneda3.ListIndex = 0
    Desdelugar = 0

    XCoti.Height = 4200
    XCoti.Left = 2360
    XCoti.Top = 1320
    XCoti.Width = 7455
    
    XCoti.Visible = True
    
    XProve.Text = ""
    XArti.Text = "  -   -   "
    XPrec.Text = ""
    XCondicion.Text = ""
    XObservaciones.Text = ""
    
    XProve.SetFocus

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
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = rstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + rstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstProveedor.Close
            
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
    
    WPrimer = DBGrid1.FirstRow
    WFila = DBGrid1.Row
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    WPorceDerechos.Text = XPorceDerechos(WLugar)
    
    
    WArticulo.SetFocus

End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "ATENCION : La fecha de la orden de compra es incorrecta"
        G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
    End If
    
    If TipoPago.ListIndex = 0 Then
        m$ = "ATENCION : Se debe informar el tipo de pago de la letra"
        G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
    End If
    
    If TipoOrden.ListIndex = 1 Then
        If TipoImpo.ListIndex = 0 Then
            m$ = "ATENCION : Se debe informar la via de transporte"
            G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
        End If
    End If
    
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 0 Then
            m$ = "ATENCION : Se debe informar la condicion de la importacion"
            G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
        End If
    End If
            
    If TipoOrden.ListIndex = 1 Then
        If Leyenda.ListIndex = 1 Or Leyenda.ListIndex = 5 Then
            If Val(Flete.Text) = 0 Then
                m$ = "ATENCION : Se debe informar el monto del flete"
                G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
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
                    Else
            m$ = "Se debe informar la paridad"
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            Call Conecta_Empresa
            Exit Sub
        End If
        
        If Val(Carpeta.Text) = 0 Then
        
            Sql1 = "Select Max(Carpeta) as [CarpetaMayor]"
            Sql2 = " FROM NroCarpeta"
            spNroCarpeta = Sql1 + Sql2
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
            If rstNroCarpeta.RecordCount > 0 Then
                rstNroCarpeta.MoveLast
                ZCarpeta = IIf(IsNull(rstNroCarpeta!CarpetaMayor), "0", rstNroCarpeta!CarpetaMayor)
                Carpeta.Text = ZCarpeta + 1
                rstNroCarpeta.Close
            End If
            
            m$ = "La carpera asignada es la " + Carpeta.Text
            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM NroCarpeta"
        ZSql = ZSql + " Where Planta = " + "'" + XEmpresa + "'"
        ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
        spNroCarpeta = ZSql
        Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
        If rstNroCarpeta.RecordCount > 0 Then
            rstNroCarpeta.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE NroCarpeta SET "
            ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "'"
            ZSql = ZSql + " Where Planta = " + "'" + XEmpresa + "'"
            ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
            spNroCarpeta = ZSql
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ZSql + "INSERT INTO NroCarpeta ("
            ZSql = ZSql + "Carpeta ,"
            ZSql = ZSql + "Planta ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Carpeta.Text + "',"
            ZSql = ZSql + "'" + XEmpresa + "',"
            ZSql = ZSql + "'" + Orden.Text + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "')"
            spNroCarpeta = ZSql
            Set rstNroCarpeta = db.OpenRecordset(spNroCarpeta, dbOpenSnapshot, dbSQLPassThrough)
        End If
            
        Call Conecta_Empresa
        
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
    
    For DA = 1 To Renglon
    
        Articulo = Vector(DA, 1)
        Cantidad = Vector(DA, 2)
    
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
        
    Next DA
        
    spOrden = "BorrarOrdenTotal " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenDynaset, dbSQLPassThrough)
        
    Renglon = 0
    ZSuma = 0
    ZBase = 0
    Erase ZVector
        
    DBGrid1.Refresh
        
    For a = 0 To 9
        
        Suma = a * 10
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
            Fecha2 = DBGrid1.Text
            If Fecha2 <> "" Then
                ZZAuxiFecha = Fecha2
            End If
                    
            DBGrid1.Col = 6
            Condicion = DBGrid1.Text
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
            WWPorceDerechos = XPorceDerechos(WLugar)
                    
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
                WFecha2 = Fecha2
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
                
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Orden SET "
                ZSql = ZSql + " Derechos = " + "'" + WWPorceDerechos + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spOrden = ZSql
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
            
    Next a
    
    If TipoOrden.ListIndex = 1 Then
    
        ZDespacho = 0
        ZBase = 0
        ZBaseII = 0
        ZBaseIII = 0
        ZBaseIV = 0
        
        ZSuma1 = 0
        ZSuma2 = 0
        ZSuma3 = 0
        ZSuma4 = 0
        ZSuma5 = 0
        ZSuma6 = 0
        ZSuma7 = 0
        
        If ZSuma <> 0 Then
    
            ZRegion = 0
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            With rstProveedor
                If rstProveedor.RecordCount > 0 Then
                    ZRegion = 0
                    ZRegion = IIf(IsNull(!Region), "0", !Region)
                End If
                rstProveedor.Close
            End With
    
            For Ciclo = 1 To Renglon
    
                XXArticulo = ZVector(Ciclo, 1)
                XXCantidad = ZVector(Ciclo, 2)
                XXPrecio = ZVector(Ciclo, 3)
                ZPorceDerechos = Val(ZVector(Ciclo, 4))
                
                If ZRegion = 0 And ZPorceDerechos = 0 Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Articulo = " + "'" + XXArticulo + "'"
                    ZSql = ZSql + " and Derechos <> 0"
                    ZSql = ZSql + " Order by Orden.FechaOrd"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        With rstOrden
                            .MoveLast
                            ZPorceDerechos = rstOrden!Derechos
                        End With
                        rstOrden.Close
                    End If
                    If ZPorceDerechos = 0 Then
                        m$ = "ATENCION : Se debe informar los derechos para " + XXArticulo
                        G% = MsgBox(m$, 48, "Ingreso de Orden de Compra")
                    End If
                End If
            
                ZImpo = Val(XXCantidad) * Val(XXPrecio)
                ZBaseII = ZBaseII + ZImpo
            
                ZSeguro = 0
                If Val(Flete.Text) <> 0 Then
                    ZPorce = ZImpo / ZSuma
                    ZFlete = Val(Flete.Text) * ZPorce
                    ZImpo = ZImpo + ZFlete
                End If
                If Leyenda.ListIndex <> 2 Then
                    ZSeguro = ZImpo * 0.01
                End If
                ZBaseIV = ZBaseIV + ZSeguro
        
                ZImpo = ZImpo + ZSeguro
                ZBaseIII = ZBaseIII + ZImpo
            
                ZDerechos = ZImpo * (ZPorceDerechos / 100)
                ZSuma1 = ZSuma1 + ZDerechos
                
                If ZRegion = 1 Then
                    ZEstadistica = 0
                        Else
                    ZEstadistica = ZImpo * 0.005
                End If
                ZSuma2 = ZSuma2 + ZEstadistica
        
                ZImpo = ZImpo + ZDerechos + ZEstadistica
        
                ZIva = ZImpo * 0.21
                ZIvaComp = ZImpo * 0.1
                ZGanancia = ZImpo * 0.03
                ZIBruto = ZImpo * 0.015
                
                ZSuma3 = ZSuma3 + ZIva
                ZSuma4 = ZSuma4 + ZIvaComp
                ZSuma5 = ZSuma5 + ZGanancia
                ZSuma6 = ZSuma6 + ZIBruto
        
                ZImpo = ZImpo + ZIva + ZIvaComp + ZGanancia + ZIBruto
                ZBase = ZBase + ZImpo
        
            Next Ciclo
            
            ZCargo = 10
            ZBase = ZBase + ZCargo - ZBaseIII
    
            ZImpoII = ZBase * ZParidad
            ZImpoIV = ZBaseIII * ZParidad
    
            ZGastos = 100
            ZHonorarios = ZImpoIV * 0.006
            ZIvaGastos = (ZGastos + ZHonorarios) * 0.21
        
            Select Case TipoImpo.ListIndex
                Case 1
                    ZViaI = 1200
                    ZViaII = 1200
                Case 2
                    ZViaI = 250
                    ZViaII = 0
                Case 3
                    ZViaI = 200 * ZParidad
                    ZViaII = 145 * ZParidad
                Case Else
            End Select
    
            ZDespachoI = ZGastos + ZHonorarios + ZIvaGastos + ZViaI + ZViaII + (Val(Flete.Text) * ZParidad)
            Rem  (ZBaseIV * ZParidad)
            ZDespacho = ZImpoII + ZDespachoI
    
        End If
        
            Else
            
        ZDespacho = 0
        
    End If
    
    ZDespacho = Int(ZDespacho)
    
    ZOrigen = Origen.Text
    ZLeyenda = Str$(Leyenda.ListIndex)
    ZPedidoImpo = PedidoImpo.Text
    ZFlete = Flete.Text
    ZFechaImpo = FechaImpo.Text
    ZOrdFechaImpo = Right$(FechaImpo.Text, 4) + Mid$(FechaImpo.Text, 4, 2) + Left$(FechaImpo.Text, 2)
    ZTipoImpo = Str$(TipoImpo.ListIndex)
    ZTipoPago = Str$(TipoPago.ListIndex)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " ImpoDespacho = " + "'" + Str$(ZDespacho) + "',"
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
    
    Rem ZZFechaLlegada = "  /  /    "
    Rem ZZPagoDespacho = "0"
    Rem ZZImpoDespacho = "0"
    Rem ZZVtoDespacho = "  /  /    "
    Rem ZZImpoLetra = "0"
    Rem ZZPagoLetra = "0"
    Rem ZZVtoLetra = "  /  /    "
    
    If ZZFechaLlegada = "  /  /    " Or Trim(ZZFechaLlegada) = "" Then
        ZZFechaLlegada = ZZAuxiFecha
    End If
    
    If TipoPago.ListIndex = 1 Then
        ZZVtoLetra = Fecha.Text
    End If
    If TipoPago.ListIndex = 2 Then
        ZZVtoLetra = ZZFechaLlegada
    End If
    ZZImpoLetra = Str$(ZSuma)
    ZZVtoDespacho = ZZFechaLlegada

    ZZOrdVtoDespacho = Right$(ZZVtoDespacho, 4) + Mid$(ZZVtoDespacho, 4, 2) + Left$(ZZVtoDespacho, 2)
    ZZOrdVtoLetra = Right$(ZZVtoLetra, 4) + Mid$(ZZVtoLetra, 4, 2) + Left$(ZZVtoLetra, 2)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Orden SET "
    ZSql = ZSql + " FechaLlegada = " + "'" + ZZFechaLlegada + "',"
    ZSql = ZSql + " PagoDespacho = " + "'" + ZZPagoDespacho + "',"
    ZSql = ZSql + " VtoDespacho = " + "'" + ZZVtoDespacho + "',"
    ZSql = ZSql + " OrdVtoDespacho = " + "'" + ZZOrdVtoDespacho + "',"
    ZSql = ZSql + " PagoLetra = " + "'" + ZZPagoLetra + "',"
    ZSql = ZSql + " ImpoLetra = " + "'" + ZZImpoLetra + "',"
    ZSql = ZSql + " VtoLetra = " + "'" + ZZVtoLetra + "',"
    ZSql = ZSql + " OrdVtoLetra = " + "'" + ZZOrdVtoLetra + "'"
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
        
    T$ = "Orden de Compra"
    m$ = "Desea enviar la O/C via email al proveedor"
    Respuesta% = MsgBox(m$, 256 + 4, T$)
    If Respuesta% = 6 Then
        Call EMail_Click
        Rem ChDir "\\PRUEBA\E\VB"
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

    DatosImpo.Visible = False

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
    ZEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    Flete.Text = ""
    
    Moneda.ListIndex = 0
    TipoOrden.ListIndex = 1
    TipoPago.ListIndex = 0
    Leyenda.ListIndex = 0
    TipoImpo.ListIndex = 0
    
    For a = 0 To 9
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 6
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
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
    ZSql = ZSql + " Where Orden.Orden < 800000"
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

Private Sub TipoOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    If TipoOrden.ListIndex = 3 Then
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden < 900000"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveLast
                If rstOrden!Orden >= 800000 Then
                    Orden.Text = Mid$(Str$(rstOrden!Orden + 1), 2, 6)
                        Else
                    Orden.Text = "800000"
                End If
            End With
            rstOrden.Close
        End If
            Else
        If Val(Orden.Text) >= 800000 Then
            ZSql = ""
            ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Orden < 800000"
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
        End If
    End If
    End If
End Sub

Private Sub TipoOrden_Click()
    If TipoOrden.ListIndex = 3 And Val(Orden.Text) < 800000 Then
        ZSql = ""
        ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden < 900000"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveLast
                If rstOrden!Orden >= 800000 Then
                    Orden.Text = Mid$(Str$(rstOrden!Orden + 1), 2, 6)
                        Else
                    Orden.Text = "800000"
                End If
            End With
            rstOrden.Close
        End If
            Else
        If TipoOrden.ListIndex <> 3 And Val(Orden.Text) >= 800000 Then
            ZSql = ""
            ZSql = ZSql + "Select Orden.Clave, Orden.Orden"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden.Orden < 800000"
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
        End If
    End If
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
            
            If TipoOrden.ListIndex = 1 Then
            
                ZRegion = 0
                ZZPorceDerechos = Val(WPorceDerechos.Text)
                spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                With rstProveedor
                    If rstProveedor.RecordCount > 0 Then
                        ZRegion = 0
                        ZRegion = IIf(IsNull(!Region), "0", !Region)
                    End If
                    rstProveedor.Close
                End With
            
                If ZRegion = 0 And ZZPorceDerechos = 0 Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                    ZSql = ZSql + " and Derechos <> 0"
                    ZSql = ZSql + " Order by Orden.FechaOrd"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        With rstOrden
                            .MoveLast
                            ZZPorceDerechos = rstOrden!Derechos
                        End With
                        rstOrden.Close
                    End If
                End If

                If ZRegion = 0 And ZZPorceDerechos = 0 Then
                    WPorceDerechos.Text = ""
                    IngreDerechos.Visible = True
                    WPorceDerechos.SetFocus
                    Exit Sub
                End If
            
                If ZRegion = 1 Then
                    WPorceDerechos.Text = ""
                End If
            
                WPorceDerechos.Text = Str$(ZZPorceDerechos)
                
            End If
        
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
        
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
            
                Else
                
            WFecha2.SetFocus
            
        End If
    End If
End Sub

Private Sub WPorceDerechos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WPorceDerechos.Text) <> 0 Then
            IngreDerechos.Visible = False
            Call Alta_Vector
            Call Ingresa_Click
            WArticulo.SetFocus
                Else
            WPorceDerechos.SetFocus
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
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Select Case Val(TipoConsulta)
                    Case 2
                        XProv1.Text = WProveedor
                        XDesProv1.Caption = rstProveedor!Nombre
                        XProv1.SetFocus
                    Case 3
                        XProve.Text = WProveedor
                        XDesProve.Caption = rstProveedor!Nombre
                        XProve.SetFocus
                    Case 4
                        XProv3.Text = WProveedor
                        XDesProv3.Caption = rstProveedor!Nombre
                        XProv3.SetFocus
                    Case Else
                        Proveedor.Text = WProveedor
                        DesProveedor.Caption = rstProveedor!Nombre
                        ZEmail = rstProveedor!email
                        Proveedor.SetFocus
                End Select
                rstProveedor.Close
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
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
                rstArticulo.Close
                    
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
    ZEmail = ""
    Carpeta.Text = ""
    Origen.Text = ""
    PedidoImpo.Text = ""
    FechaImpo.Text = "  /  /    "
    Flete.Text = ""
    
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
    TipoOrden.AddItem "Envases"

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
    ZSql = ZSql + " Where Orden.Orden < 800000"
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
    
    For a = 0 To 9
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 6
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
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
                DBGrid1.Text = rstOrden!Fecha2
                
                DBGrid1.Col = 6
                DBGrid1.Text = rstOrden!Condicion
                
                WPrimer = DBGrid1.FirstRow
                WFila = DBGrid1.Row
                WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
                XPorceDerechos(WLugar) = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                
                If rstOrden!Recibida > 0 Then
                    Graba.Enabled = False
                End If
                
                Rem ZMarca = IIf(IsNull(rstOrden!Marca), "", rstOrden!Marca)
                Rem If ZMarca = "X" Then
                Rem     Graba.Enabled = False
                Rem End If
                
                Vector(Renglon, 1) = Auxi1
            
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstOrden.Close
    
    If Graba.Enabled = False Then
        m$ = "La orden de compra no podra ser actualizada ya que posee productos que fueron cumplidos en forma total o parcial, o cargados datos adicionales referentes a la importacion"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To WRenglon
    
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
    Next DA
    
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
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
            
            XPorceDerechos(WLugar) = WPorceDerechos.Text
            
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
            
            WPrimer = DBGrid1.FirstRow
            WFila = DBGrid1.Row
            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
            
            XPorceDerechos(WLugar) = WPorceDerechos.Text
            
            Rem DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
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
            Flete.Text = IIf(IsNull(rstOrden!Flete), "", rstOrden!Flete)
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Proveedor.Text = rstProveedor!Proveedor
                DesProveedor.Caption = rstProveedor!Nombre
                ZEmail = rstProveedor!email
                rstProveedor.Close
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                Proveedor.Text = rstProveedor!Proveedor
                DesProveedor.Caption = rstProveedor!Nombre
                ZEmail = rstProveedor!email
                rstProveedor.Close
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
    
    a = 1

End Sub


Sub Impresion()

        Open "lpt1" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        For Ci = 1 To 2

        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        If TipoOrden.ListIndex = 1 Then
            Print #1, Tab(30); "(IMPORTACION)";
        End If
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 40);
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        
        If Leyenda.ListIndex <> 0 Then
            Print #1, Tab(50); "("; Leyenda.Text; ")":
        End If
        
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion  |Canti.|   Precio  |1ra Fec.  |Ul.Fecha  |Cond. Pago|"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        
        For a = 0 To 9
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                    
                If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
                
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = Articulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                            Else
                        DBGrid1.Col = 1
                        WDescripcion = DBGrid1.Text
                    End If
                        
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Fecha1 = DBGrid1.Text
                    
                    DBGrid1.Col = 5
                    Fecha2 = DBGrid1.Text
                    
                    DBGrid1.Col = 6
                    Condicion = DBGrid1.Text

                    WCantidad = WCantidad + 1

                    Print #1, Tab(1); "|"; Articulo;
                    Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                    Print #1, Tab(28); "|"; Alinea("##,###", Str$(Cantidad));
                    Select Case Moneda.ListIndex
                        Case 0
                            Print #1, Tab(35); "|U$S"; Alinea("#,###.##", Str$(Precio));
                        Case Else
                            Print #1, Tab(35); "| $ "; Alinea("#,###.##", Str$(Precio));
                    End Select
                    Print #1, Tab(47); "|"; Fecha1;
                    Print #1, Tab(58); "|"; Fecha2;
                    Print #1, Tab(69); "|"; Left$(Condicion, 10);
                    Print #1, Tab(80); "|"

                    Valor = Valor + (Cantidad * Precio)

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |               |      |           |          |          |          |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        If Moneda.ListIndex = 0 Then
            Print #1, "|          Valor total de la orden : U$S "; Alinea("#####.##", Str$(Valor));
                Else
            Print #1, "|          Valor total de la orden :   $ "; Alinea("#####.##", Str$(Valor));
        End If
        Print #1, Tab(80); "|"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Next Ci

ImpreCopia:

        WCantidad = 0
        
        ' Copia 2

        Rem Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(60); "Remito :..........";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        If TipoOrden.ListIndex = 1 Then
            Print #1, Tab(30); "(IMPORTACION)";
        End If
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 20);
        Print #1, Tab(60); "Informe :.........";
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
        Print #1, "--------------------------------------------------------------------------------"

        Cantidad = 0
        Valor = 0
        
        For a = 0 To 9
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                
                If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
                
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = Articulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                            Else
                        DBGrid1.Col = 1
                        WDescripcion = DBGrid1.Text
                    End If
                    
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Fecha1 = DBGrid1.Text
                    
                    DBGrid1.Col = 5
                    Fecha2 = DBGrid1.Text
                    
                    DBGrid1.Col = 6
                    Condicion = DBGrid1.Text
                
                    WUbicacion = ""
                
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WUbicacion = rstArticulo!Deposito
                        rstArticulo.Close
                    End If
                
                    WCantidad = WCantidad + 2

                    Print #1, Tab(1); "|"; Articulo;
                    Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
                    Print #1, Tab(41); "|"; Alinea("###,###", Str$(Cantidad));
                    Print #1, Tab(50); "|"; Fecha1;
                    Print #1, Tab(61); "|"; Fecha2;
                    Print #1, Tab(72); "|";
                    Print #1, Tab(80); "|"
                        
                    Print #1, Tab(1); "|";
                    Print #1, Tab(12); "|"; WUbicacion;
                    Print #1, Tab(50); "|";
                    Print #1, Tab(61); "|";
                    Print #1, Tab(72); "|";
                    Print #1, Tab(80); "|"

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |                            |        |          |          |       |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

    Close #1

End Sub

Private Sub Primera_Click()

        Open "lpt1" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
    
        '  Copia 1
        
        Print #1, Chr$(18)
        Print #1, ""
        Print #1, ""

        Print #1, Tab(1); "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|                                                                              |"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        If TipoOrden.ListIndex = 1 Then
            Print #1, Tab(30); "(IMPORTACION)";
        End If
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 40);
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        
        If Leyenda.ListIndex <> 0 Then
            Print #1, Tab(50); "("; Leyenda.Text; ")";
        End If
        
        Print #1, Tab(80); "|"
        
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |  Descripcion  |Canti.|   Precio  |1ra Fec.  |Ul.Fecha  |Cond. Pago|"
        Print #1, "--------------------------------------------------------------------------------"

        WCantidad = 0
        Valor = 0
        
        For a = 0 To 9
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                    
                If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
                
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = Articulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                            Else
                        DBGrid1.Col = 1
                        WDescripcion = DBGrid1.Text
                    End If
                    
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Fecha1 = DBGrid1.Text
                    
                    DBGrid1.Col = 5
                    Fecha2 = DBGrid1.Text
                    
                    DBGrid1.Col = 6
                    Condicion = DBGrid1.Text

                    WCantidad = WCantidad + 1

                    Print #1, Tab(1); "|"; Articulo;
                    Print #1, Tab(12); "|"; Left$(WDescripcion, 15);
                    Print #1, Tab(28); "|"; Alinea("##,###", Str$(Cantidad));
                    Select Case Moneda.ListIndex
                        Case 0
                            Print #1, Tab(35); "|U$S"; Alinea("#,###.##", Str$(Precio));
                        Case Else
                            Print #1, Tab(35); "| $ "; Alinea("#,###.##", Str$(Precio));
                    End Select
                    Print #1, Tab(47); "|"; Fecha1;
                    Print #1, Tab(58); "|"; Fecha2;
                    Print #1, Tab(69); "|"; Left$(Condicion, 10);
                    Print #1, Tab(80); "|"

                    Valor = Valor + (Cantidad * Precio)

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |               |      |           |          |          |          |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        If Moneda.ListIndex = 0 Then
            Print #1, "|          Valor total de la orden : U$S "; Alinea("#####.##", Str$(Valor));
                Else
            Print #1, "|          Valor total de la orden :   $ "; Alinea("#####.##", Str$(Valor));
        End If
        Print #1, Tab(80); "|"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

        Close #1

 End Sub

Private Sub Tercera_Click()

        Open "lpt1" For Output As #1
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
           .Seek "=", Claveven$
           If .NoMatch = False Then
               Impretit = !Nombre
                   Else
               Impretit = ""
           End If
        End With
    
        WCantidad = 0
        
        ' Copia 2

        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        
        Print #1, Tab(1); "|";
        Print #1, "Empresa : "; Impretit;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(60); "Remito :..........";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Orden.....: ";
        Print #1, Tab(20); Alinea("######", Orden.Text);
        Print #1, Tab(50); "Fecha : "; Fecha.Text;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Proveedor...:"; Tab(20); Proveedor.Text;
        Print #1, Tab(35); Left$(DesProveedor.Caption, 20);
        Print #1, Tab(60); "Informe :.........";
        Print #1, Tab(80); "|"
        
        If Val(Carpeta.Text) <> 0 Then
            Print #1, Tab(1); "|";
            Print #1, Tab(5); "Carpeta.....:"; Tab(20); Carpeta.Text;
                Else
            Print #1, Tab(1); "|";
        End If
        Print #1, Tab(80); "|"
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "|Producto  |        Descripcion         |  Canti.|1ra Fec.  |Ul.Fecha  |F.Recep|"
        Print #1, "--------------------------------------------------------------------------------"

        Cantidad = 0
        Valor = 0
        
        For a = 0 To 9
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = UCase(DBGrid1.Text)
                
                If Left$(Articulo, 2) <> "" And Left$(Articulo, 2) <> Space$(2) Then
                
                    XProveedor = Proveedor.Text
                    Call Ceros(XProveedor, 11)
                    ClaveMarcas = Articulo + XProveedor
                    spMarcas = "ConsultaMarcas " + "'" + ClaveMarcas + "'"
                    Set rstMarcas = db.OpenRecordset(spMarcas, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMarcas.RecordCount > 0 Then
                        WDescripcion = rstMarcas!Descripcion
                        rstMarcas.Close
                            Else
                        DBGrid1.Col = 1
                        WDescripcion = DBGrid1.Text
                    End If
                    
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 4
                    Fecha1 = DBGrid1.Text
                    
                    DBGrid1.Col = 5
                    Fecha2 = DBGrid1.Text
                    
                    DBGrid1.Col = 6
                    Condicion = DBGrid1.Text
                
                    WUbicacion = ""
                
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WUbicacion = rstArticulo!Deposito
                        rstArticulo.Close
                    End If

                    WCantidad = WCantidad + 2

                    Print #1, Tab(1); "|"; Articulo;
                    Print #1, Tab(12); "|"; Left$(WDescripcion, 28);
                    Print #1, Tab(41); "|"; Alinea("###,###", Str$(Cantidad));
                    Print #1, Tab(50); "|"; Fecha1;
                    Print #1, Tab(61); "|"; Fecha2;
                    Print #1, Tab(72); "|";
                    Print #1, Tab(80); "|"
                        
                    Print #1, Tab(1); "|";
                    Print #1, Tab(12); "|"; WUbicacion;
                    Print #1, Tab(50); "|";
                    Print #1, Tab(61); "|";
                    Print #1, Tab(72); "|";
                    Print #1, Tab(80); "|"

                End If
                                        
            Next iRow
        Next a

        For Ciclo = WCantidad To 15
            Print #1, "|          |                            |        |          |          |       |"
        Next Ciclo

        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""

    Close #1
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Orden.SetFocus

 End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)

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
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstProveedor.RecordCount > 0 Then
            With rstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        DA = Len(rstProveedor!Nombre) - WEspacios
                
                        For aa = 1 To DA
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                Auxi = Str$(rstProveedor!Proveedor)
                                Call Ceros(Auxi, 11)
                                IngresaItem = Auxi + "    " + rstProveedor!Nombre
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
            rstProveedor.Close
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
            
                            DA = Len(rstArticulo!Descripcion) - WEspacios
                
                            For Aaa = 1 To DA
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

Private Sub XAcepta1_Click()

    XEmpresa = WEmpresa
        
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    DA = 0
    With rstLiscot
        .Index = "Clave"
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
    
    WAno = Right$(Date$, 4)
    WDia = Mid$(Date$, 4, 2)
    WMes = Left$(Date$, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            AA1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = rstCambios!Cambio
            rstCambios.Close
        End With
            Else
        Paridad = 1
    End If
    
    XParam = "'" + XProv1.Text + "','" _
            + XProv1.Text + "'"
    
    spCotiza = "ListaCotizaProveedorDesdeHasta" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    
    Pasa = 0
    Canti = 0
    
    If rstCotiza.RecordCount > 0 Then
            
        With rstCotiza
            .MoveFirst
    
            Do
            
                If .EOF = True Then
                    Exit Do
                End If

                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                Select Case Moneda1.ListIndex
                    Case 0
                        If !Moneda = 0 Then
                            WPrecio = !Precio
                                Else
                            WPrecio = !Precio / Paridad
                        End If
                    Case Else
                        If !Moneda = 1 Then
                            WPrecio = !Precio
                                Else
                            WPrecio = !Precio * Paridad
                        End If
                End Select
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                
                        For DA = 1 To 9
                        
                            If XVector(DA, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = XVector(DA, 1)
                                !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                !Precio = Val(XVector(DA, 2))
                                !Condicion = XVector(DA, 3)
                                !Observaciones = XVector(DA, 4)
                                !Clave = !Proveedor + !Articulo
                                !Orden = 0
                                .Update
                            End If
                            
                        Next DA
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For DA = 1 To 2
                        XVector(DA, 1) = XVector(DA + 1, 1)
                        XVector(DA, 2) = XVector(DA + 1, 2)
                        XVector(DA, 3) = XVector(DA + 1, 3)
                        XVector(DA, 4) = XVector(DA + 1, 4)
                    Next DA
                    Canti = 3
                End If
                
                XVector(Canti, 1) = !Fecha
                XVector(Canti, 2) = Str$(WPrecio)
                XVector(Canti, 3) = !Condicion
                XVector(Canti, 4) = !Observaciones
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
        End With
    End If
    
    If Pasa <> 0 Then
        With rstLiscot
                
            For DA = 1 To 3
                    
                If XVector(DA, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = XVector(DA, 1)
                    !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    !Precio = Val(XVector(DA, 2))
                    !Condicion = XVector(DA, 3)
                    !Observaciones = XVector(DA, 4)
                    !Clave = !Proveedor + !Articulo
                    .Update
                End If
                
            Next DA
                        
        End With
    End If
    
    DA = 0
    With rstLiscot
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WDescriProveedor = ""
                WArticulo = !Articulo
                WDescriArticulo = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                
                    WDescriProveedor = rstProveedor!Nombre
                    
                    ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
                    ZCategoriaII = IIf(IsNull(rstProveedor!CategoriaII), "0", rstProveedor!CategoriaII)
                    
                    WCategoriaI = ""
                    WCategoriaII = ""
        
                    If ZCategoriaI = 1 Then
                        WCategoriaI = "A"
                            Else
                        If ZCategoriaI = 2 Then
                            WCategoriaI = "B"
                                Else
                            If ZCategoriaI = 3 Then
                                WCategoriaI = "C"
                                    Else
                                If ZCategoriaI = 4 Then
                                    WCategoriaI = "E"
                                End If
                            End If
                        End If
                    End If
                    
                    WCategoriaII = "S/C"
                    If ZCategoriaII = 1 Then
                        WCategoriaII = "Muy Bueno"
                            Else
                        If ZCategoriaII = 2 Then
                            WCategoriaII = "Bueno"
                                Else
                            If ZCategoriaII = 3 Then
                                WCategoriaII = "Regular"
                                    Else
                                If ZCategoriaII = 4 Then
                                    WCategoriaII = "Malo"
                                End If
                            End If
                        End If
                    End If
                    
                    If WCategoriaI <> "" And WCategoriaII <> "" Then
                        WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                    End If
                    
                    rstProveedor.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DescriProveedor = WDescriProveedor
                !DescriArticulo = WDescriArticulo
                
                Select Case Moneda1.ListIndex
                    Case 0
                        !Titulo = "(En Dolares)"
                    Case Else
                        !Titulo = "(En Pesos)"
                End Select
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cotizaciones por Proveedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Listcot.proveedor} in " + Chr$(34) + XProv1.Text + Chr$(34) + " to " + Chr$(34) + XProv1.Text + Chr$(34)
    
    Listado.ReportFileName = "WCotprv.rpt"
    Rem Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Destination = 0
    Listado.Action = 1
    
    Call Conecta_Empresa
    
End Sub

Private Sub XAcepta2_Click()

    XEmpresa = WEmpresa
        
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    XArt2.Text = UCase(XArt2.Text)

    DA = 0
    With rstLiscot
        .Index = "Clave"
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
    
    WAno = Right$(Date$, 4)
    WDia = Mid$(Date$, 4, 2)
    WMes = Left$(Date$, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            AA1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = rstCambios!Cambio
            rstCambios.Close
        End With
            Else
        Paridad = 1
    End If
    
    
    Pasa = 0
    Canti = 0
    
    XParam = "'" + XArt2.Text + "','" _
            + XArt2.Text + "'"
    
    spCotiza = "ListaCotizaArticuloDesdeHasta" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
            
    With rstCotiza
    
            .MoveFirst
            
            Do
            
                WCotiza = !Cotiza
                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                Select Case Moneda2.ListIndex
                    Case 0
                        If !Moneda = 0 Then
                            WPrecio = !Precio
                                Else
                            WPrecio = !Precio / Paridad
                        End If
                    Case Else
                        If !Moneda = 1 Then
                            WPrecio = !Precio
                                Else
                            WPrecio = !Precio * Paridad
                        End If
                End Select
                
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                    
                        Rem If Val(XVector(3, 2)) <> 0 Then
                        Rem     WAuxi = Int(Val(XVector(3, 2)) * 100)
                        Rem             Else
                        Rem     If Val(XVector(2, 2)) <> 0 Then
                        Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
                        Rem             Else
                        Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
                        Rem     End If
                        Rem End If
                        Rem
                        Rem Call Ceros(WAuxi, 9)
                        
                        If XVector(3, 5) <> "" Then
                            WAuxi = XVector(3, 5)
                                    Else
                            If XVector(2, 5) <> "" Then
                                WAuxi = XVector(2, 5)
                                    Else
                                WAuxi = XVector(1, 5)
                            End If
                        End If
                        WAuxi = Str$(Val(WAuxi) - 90000000)
                        Call Ceros(WAuxi, 9)
                    
                        For DA = 1 To 9
                        
                            If XVector(DA, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = XVector(DA, 1)
                                !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                !Precio = Val(XVector(DA, 2))
                                !Condicion = XVector(DA, 3)
                                !Observaciones = XVector(DA, 4)
                                !Clave = !Proveedor + !Articulo
                                !Orden = WAuxi + !Proveedor
                                .Update
                            End If
                            
                        Next DA
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase XVector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For DA = 1 To 2
                        XVector(DA, 1) = XVector(DA + 1, 1)
                        XVector(DA, 2) = XVector(DA + 1, 2)
                        XVector(DA, 3) = XVector(DA + 1, 3)
                        XVector(DA, 4) = XVector(DA + 1, 4)
                        XVector(DA, 5) = XVector(DA + 1, 5)
                    Next DA
                    Canti = 3
                End If
                
                XVector(Canti, 1) = !Fecha
                XVector(Canti, 2) = Str$(WPrecio)
                XVector(Canti, 3) = !Condicion
                XVector(Canti, 4) = !Observaciones
                XVector(Canti, 5) = !FechaOrd
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    rstCotiza.Close
    
    End If
    
    If Pasa <> 0 Then
        With rstLiscot
        
            Rem If Val(XVector(3, 2)) <> 0 Then
            Rem     WAuxi = Int(Val(XVector(3, 2)) * 100)
            Rem             Else
            Rem     If Val(XVector(2, 2)) <> 0 Then
            Rem         WAuxi = Int(Val(XVector(2, 2)) * 100)
            Rem             Else
            Rem         WAuxi = Int(Val(XVector(1, 2)) * 100)
            Rem     End If
            Rem End If
            Rem
            Rem Call Ceros(WAuxi, 9)
            
            If XVector(3, 5) <> "" Then
                WAuxi = XVector(3, 5)
                        Else
                If XVector(2, 5) <> "" Then
                    WAuxi = XVector(2, 5)
                        Else
                    WAuxi = XVector(1, 5)
                End If
            End If
            WAuxi = Str$(Val(WAuxi) - 90000000)
            Call Ceros(WAuxi, 9)
                
            For DA = 1 To 9
                    
                If XVector(DA, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = XVector(DA, 1)
                    !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    !Precio = Val(XVector(DA, 2))
                    !Condicion = XVector(DA, 3)
                    !Observaciones = XVector(DA, 4)
                    !Clave = !Proveedor + !Articulo
                    !Orden = WAuxi + !Proveedor
                    .Update
                End If
                
            Next DA
                        
        End With
    End If
    
    DA = 0
    With rstLiscot
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WDescriProveedor = ""
                WArticulo = !Articulo
                WDescriArticulo = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                
                    WDescriProveedor = rstProveedor!Nombre
                    
                    ZCategoriaI = IIf(IsNull(rstProveedor!CategoriaI), "0", rstProveedor!CategoriaI)
                    ZCategoriaII = IIf(IsNull(rstProveedor!CategoriaII), "0", rstProveedor!CategoriaII)
                    
                    WCategoriaI = ""
                    WCategoriaII = ""
        
                    If ZCategoriaI = 1 Then
                        WCategoriaI = "A"
                            Else
                        If ZCategoriaI = 2 Then
                            WCategoriaI = "B"
                                Else
                            If ZCategoriaI = 3 Then
                                WCategoriaI = "C"
                                    Else
                                If ZCategoriaI = 4 Then
                                    WCategoriaI = "E"
                                End If
                            End If
                        End If
                    End If
                    
                    WCategoriaII = "S/C"
                    If ZCategoriaII = 1 Then
                        WCategoriaII = "Muy Bueno"
                            Else
                        If ZCategoriaII = 2 Then
                            WCategoriaII = "Bueno"
                                Else
                            If ZCategoriaII = 3 Then
                                WCategoriaII = "Regular"
                                    Else
                                If ZCategoriaII = 4 Then
                                    WCategoriaII = "Malo"
                                End If
                            End If
                        End If
                    End If
                    
                    If WCategoriaI <> "" And WCategoriaII <> "" Then
                        WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                    End If
                    
                    rstProveedor.Close
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DescriProveedor = WDescriProveedor
                !DescriArticulo = WDescriArticulo
                
                Select Case Moneda2.ListIndex
                    Case 0
                        !Titulo = "(En Dolares)"
                    Case Else
                        !Titulo = "(En Pesos)"
                End Select
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cotizaciones por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Listcot.Articulo} in " + Chr$(34) + XArt2.Text + Chr$(34) + " to " + Chr$(34) + XArt2.Text + Chr$(34)
   
    Listado.Destination = 0
    Listado.ReportFileName = "WCotart.rpt"
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
    Call Conecta_Empresa

End Sub

Private Sub XAcepta3_Click()

    On Error GoTo WError
    
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

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With

    DA = ""
    With rstImpCtaCtePrv
        .Index = "ClaveImpre"
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
    
    XParam = "'" + XProv3.Text + "','" _
                 + XProv3.Text + "'"
    spCtaprv = "ListaCtaprvDesdeHasta " + XParam
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
    
    With RstCtaPrv
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                XProveedor = !Proveedor
                XLetra = !Letra
                XTipo = !Tipo
                XPunto = !Punto
                XNumero = !Numero
                XFecha = !Fecha
                XEstado = !Estado
                Xvencimiento = !Vencimiento
                XVencimiento1 = !Vencimiento1
                XNroInterno = !NroInterno
                XTotal = !Total
                XSaldo = !Saldo
                XClave = !Clave
                XOrdFecha = !OrdFecha
                XOrdVencimiento = !OrdVencimiento
                XImpre = !Impre
                
                With rstImpCtaCtePrv
                
                    .Index = "CtaCte"
                    .Seek "=", XClave
                    If .NoMatch Then
                        .AddNew
                        !Proveedor = XProveedor
                        !Letra = XLetra
                        !Tipo = XTipo
                        !Punto = XPunto
                        !Numero = XNumero
                        !Fecha = XFecha
                        !Estado = XEstado
                        !Vencimiento = Xvencimiento
                        !Vencimiento1 = XVencimiento1
                        !NroInterno = XNroInterno
                        !Total = XTotal
                        !Saldo = XSaldo
                        !Clave = XClave
                        !OrdFecha = XOrdFecha
                        !OrdVencimiento = XOrdVencimiento
                        !Impre = XImpre
                        !Titulo = WTitulo
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstCtaPrv.Close
    
    End If
    
    
    Pasa = 0
    Acumula = 0

    With rstImpCtaCtePrv
            .Index = "ClaveImpre"
            .MoveFirst
            Do
                Rem If !Proveedor > Hasta.Text Then
                Rem    Exit Do
                Rem End If
                If Pasa = 0 Then
                    Pasa = 1
                    Acumula = 0
                    corte = !Proveedor
                End If
                If corte <> !Proveedor Then
                    Acumula = 0
                    corte = !Proveedor
                End If
                .Edit
                !SaldoList = 0
                If !Proveedor >= XProv3.Text And !Proveedor <= XProv3.Text Then
                    WSaldo = !Saldo
                    Call Redondeo(WSaldo)
                    !SaldoList = WSaldo
                    Acumula = Acumula + WSaldo
                    !Acumulado = Acumula
                End If
                
                WProveedor = !Proveedor
                WNombre = ""
                WCheque = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    WNombre = rstProveedor!Nombre
                    WCheque = rstProveedor!NombreCheque
                    rstProveedor.Close
                End If
                
                !Nombre = WNombre
                !Cheque = WCheque
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Listado.GroupSelectionFormula = "{CtaCtePrv.Proveedor} in " + Chr$(34) + XProv3.Text + Chr$(34) + " to " + Chr$(34) + XProv3.Text + Chr$(34) + " and {CtaCtePrv.Saldolist} <> 0.0"
    Listado.Destination = 0
    Listado.DataFiles(0) = XEmpresa + "Auxi.mdb"
    Listado.ReportFileName = "wccprv.rpt"
    
    Listado.Action = 1
     
    Call Conecta_Empresa
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub XArt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If XArt2.Text <> "  -   -   " Then
            XArt2.Text = UCase(XArt2.Text)
            spArticulo = "ConsultaArticulo " + "'" + XArt2.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                XDesArt2.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                Call XAcepta2_Click
            End If
                Else
            TipoConsulta = "2"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 1
            Call Opcion_Click
        End If
    End If
End Sub

Private Sub XCancela_Click()
    XCoti.Visible = False
End Sub

Private Sub XCancela1_Click()
    XCotPrv.Visible = False
End Sub

Private Sub XCancela2_Click()
    XCotart.Visible = False
End Sub

Private Sub XCancela3_Click()
    XCc.Visible = False
End Sub

Private Sub XConsulta1_Click()
    TipoConsulta = "2"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Call Opcion_Click
    Ayuda.SetFocus
End Sub

Private Sub XConsulta2_Click()
    TipoConsulta = "2"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    Call Opcion_Click
End Sub

Private Sub XConsulta3_Click()
    TipoConsulta = "4"
    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Call Opcion_Click
    Ayuda.SetFocus
End Sub


Private Sub XProv1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProv1.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProv1.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProv1.Text = rstProveedor!Proveedor
                    XDesProv1.Caption = rstProveedor!Nombre
                    Call XAcepta1_Click
            End If
                Else
            TipoConsulta = "2"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)

End Sub

Private Sub XProv3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProv3.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProv3.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProv3.Text = rstProveedor!Proveedor
                    XDesProv3.Caption = rstProveedor!Nombre
                    Call XAcepta3_Click
            End If
                Else
            TipoConsulta = "4"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)

End Sub


Private Sub XProve_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(XProve.Text) <> 0 Then
            spProveedor = "Consultaproveedores " + "'" + XProve.Text + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    XProve.Text = rstProveedor!Proveedor
                    XDesProve.Caption = rstProveedor!Nombre
                    XArti.SetFocus
                        Else
                    XProve.SetFocus
            End If
                Else
            TipoConsulta = "3"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 0
            Call Opcion_Click
            Ayuda.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub XArti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If XArti.Text <> "  -   -   " Then
            XArti.Text = UCase(XArti.Text)
            spArticulo = "ConsultaArticulo " + "'" + XArti.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                XDesArti.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                XPrec.SetFocus
                    Else
                XArti.SetFocus
            End If
                Else
            TipoConsulta = "3"
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Articulos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 1
            Call Opcion_Click
        End If
    End If

End Sub

Private Sub XPrec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XCondicion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub XCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XObservaciones.SetFocus
    End If

End Sub

Private Sub XObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        XProve.SetFocus
    End If

End Sub

