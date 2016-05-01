VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDevolPesos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Devolucion de Mercaderia"
   ClientHeight    =   8355
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   11715
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8355
   ScaleWidth      =   11715
   Visible         =   0   'False
   Begin VB.TextBox factu 
      Height          =   285
      Left            =   3360
      TabIndex        =   61
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox vtocae 
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
      Left            =   9960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   60
      Text            =   " "
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Cae 
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
      Left            =   9960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   57
      Text            =   " "
      Top             =   6720
      Width           =   1695
   End
   Begin VB.ComboBox ReteIb 
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
      ItemData        =   "prgdevolPesos.frx":0000
      Left            =   8040
      List            =   "prgdevolPesos.frx":0002
      TabIndex        =   53
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox Planta 
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
      ItemData        =   "prgdevolPesos.frx":0004
      Left            =   8040
      List            =   "prgdevolPesos.frx":0006
      TabIndex        =   50
      Top             =   120
      Width           =   3015
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
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   48
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton ImpreII 
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
      Height          =   615
      Left            =   10560
      TabIndex        =   47
      Top             =   5160
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
      Left            =   120
      TabIndex        =   41
      Text            =   " "
      Top             =   5880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta de Datos"
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
      Left            =   10560
      TabIndex        =   36
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglones"
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
      Left            =   10560
      TabIndex        =   35
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   9735
      Begin VB.TextBox WEntrada 
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
         Left            =   6720
         MaxLength       =   8
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WTipopro 
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
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox WLote 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSMask.MaskEdBox WArticulo 
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
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
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   30
         Text            =   " "
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
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   240
         Width           =   2655
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
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Paridad 
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
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
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
      Left            =   10560
      TabIndex        =   26
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   6360
      TabIndex        =   17
      Top             =   5760
      Width           =   2655
      Begin VB.Label Label17 
         Caption         =   "IB Ciudad"
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
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label ImpoIbCiudad 
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
         Left            =   1200
         TabIndex        =   55
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label ImpoIbTucu 
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
         Left            =   1200
         TabIndex        =   52
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "IB Tucu."
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
         TabIndex        =   51
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "IB Bs.As."
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
         TabIndex        =   46
         Top             =   840
         Width           =   855
      End
      Begin VB.Label ImpoIb 
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
         Left            =   1200
         TabIndex        =   45
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Interes"
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
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Descuento"
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
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Interes 
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
         Left            =   1200
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Dto 
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
         Left            =   1200
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Total 
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
         Left            =   1200
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Iva2 
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
         Left            =   1200
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Iva1 
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
         Left            =   1200
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Neto 
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
         Left            =   1200
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
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
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva 10.5%"
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
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Iva 21%"
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
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Neto"
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
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6480
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Left            =   10560
      TabIndex        =   16
      Top             =   5880
      Width           =   1095
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
      Height          =   1230
      Left            =   1800
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
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
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6360
      TabIndex        =   9
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
   Begin VB.TextBox Numero 
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
      MaxLength       =   8
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
      Height          =   570
      Left            =   10560
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
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
      Height          =   570
      Left            =   10560
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
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
      Height          =   570
      Left            =   10560
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "prgdevolPesos.frx":0008
      TabIndex        =   3
      Top             =   1200
      Width           =   10335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   0
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
      Height          =   1620
      ItemData        =   "prgdevolPesos.frx":09E6
      Left            =   120
      List            =   "prgdevolPesos.frx":09ED
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label16 
      Caption         =   "Vto.Cae"
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
      Left            =   9120
      TabIndex        =   59
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Cae"
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
      Left            =   9120
      TabIndex        =   58
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Ret I.B."
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
      Left            =   6600
      TabIndex        =   54
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Nro. Devolucion"
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
      Left            =   8040
      TabIndex        =   49
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Paridad"
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
      Left            =   3840
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
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
      Top             =   840
      Width           =   1575
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
      Left            =   3480
      TabIndex        =   12
      Top             =   480
      Width           =   4455
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
      TabIndex        =   10
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
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Devolucion"
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
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "PrgDevolPesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 8 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WFecha As String
Private Wvencimiento As String
Private WVencimiento1 As String
Private WPago1 As Integer
Private WPago2 As Integer
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WCodIva As String
Private parcial As Double
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WImporte As Double
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private WDirentrega As String
Private Auxiliar(100, 10) As String
Dim IngreVector(1000, 3) As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPreciosMp  As Recordset
Dim spPreciosMp As String
Dim rstLiberaTerminado  As Recordset
Dim spLiberaTerminado As String
Dim XParam As String
Dim Compara As Double
Dim ZZZNumero As String

Private WCodIb As Integer
Private WCodIbTucu As Integer
Private WCodIbCiudad As Integer

Private WImpoIb As Double
Private WImpoIbTucu As Double
Private WImpoIbCiudad As Double
Private WPorceCm05Tucu As Double

Private WImpoPorceIb As Double
Private WImpoPorceIbTucu As Double
Private WImpoPorceIbCiudad As Double

Private WAdicional As Double
Private ZAdicional As String

Dim VectorCosto(100, 3) As String
Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZZGrabaFactura As String
Dim ZZImpreBarraI As String
Dim ZZImpreBarraII As String
Dim ZZVector(100, 5) As String
Dim ZZImpreNumero As String


Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WPlazo1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPlazo2 = rstPago!Plazo
    End If
    
    Call Calcula_vencimiento(WFecha, WPlazo2, WVencimiento1)

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
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = ""
    
    DBGrid1.Col = 7
    DBGrid1.Text = ""
    
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLinea.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    
    WArticulo.SetFocus

End Sub

Private Sub Command1_Click()
    Call Impresion
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos"

     Opcion.Visible = True
     
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    If Val(Wempresa) = 1 Then
        OPEN_FILE_Ctacte8
        OPEN_FILE_Numero8
        OPEN_FILE_Esta8
    End If
End Sub


Private Sub ImpreII_Click()
    Rem Call Impresion
    Call ImpresionFe
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
            spClientes = "ListaClienteConsulta"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                With rstClientes
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstClientes!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstClientes.Close
            End If
            
        Case 1
            Erase IngreVector
            EntraVector = 0
    
            spPreciosMp = "ListaPreciosClienteMp " + "'" + Cliente.Text + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
    
                With rstPreciosMp
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Cliente.Text = rstPreciosMp!Cliente Then
                                WArticulo = Left$(rstPreciosMp!Articulo, 3) + "00" + Right$(rstPreciosMp!Articulo, 7)
                                EntraVector = EntraVector + 1
                                IngreVector(EntraVector, 1) = WArticulo
                                IngreVector(EntraVector, 2) = rstPreciosMp!Cliente
                                IngreVector(EntraVector, 3) = rstPreciosMp!Articulo
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPreciosMp.Close
            End If
    
            For CicloVector = 1 To EntraVector
        
                WTerminado = IngreVector(CicloVector, 1)
                WCliente = IngreVector(CicloVector, 2)
                WArti = IngreVector(CicloVector, 3)
                WDescripcion = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
                IngresaItem = WTerminado + "  " + WDescripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = WCliente + WArti
                WIndice.AddItem IngresaItem
            
            Next CicloVector
        
            spPrecios = "ListaPrecios"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                With rstPrecios
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Cliente.Text = rstPrecios!Cliente Then
                                If rstPrecios!Precio <> "" Then
                                    IngresaItem = rstPrecios!Terminado + "   " + rstPrecios!Descripcion
                                        Else
                                    IngresaItem = rstPrecios!Terminado + "   " + rstPrecios!Descripcion
                                End If
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstPrecios!Cliente + rstPrecios!Terminado
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPrecios.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub DBGrid1_GotFocus()
    
    WCol = DBGrid1.Col
    WRow = DBGrid1.Row
    
    DBGrid1.Col = WCol
    DBGrid1.Row = WRow
    
    DBGrid1.Col = 0
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
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
    If Val(DBGrid1.Text) <> 0 Then
        WEntrada.Text = DBGrid1.Text
            Else
        WEntrada.Text = ""
    End If
    
    DBGrid1.Col = 5
    WTipopro.Text = DBGrid1.Text
    
    DBGrid1.Col = 6
    WLote.Text = DBGrid1.Text
    
    DBGrid1.Col = 7
    If Len(DBGrid1.Text) = 12 Then
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
            Else
        WArticulo.Text = "  -     -   "
        WLinea.Text = ""
    End If
    
    WArticulo.SetFocus
    
    If Fecha.Text = "  /  /    " Or Cliente.Text = "" Then
         Numero.SetFocus
    End If

End Sub

Private Sub Calcula_Click()

    WNeto = 0

    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
            
            DBGrid1.Col = 2
            
       Rem by nan  20 9- 2013
         Rem   Cantidad = Val(DBGrid1.Text)
            
           
            
            Cantidad = Val(IIf((DBGrid1.Text) = "", 0, (DBGrid1.Text)))
           
            
            
            
            
            
            If Cantidad <> 0 Then
                WNeto = WNeto + (Cantidad * Precio)
            End If
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    Rem If Val(Paridad.Text) <> 0 Then
    Rem    WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    WImpoDto = 0
    WImpoInteres = 0
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    If WTasa <> 0 Then
        WImpoInteres = (WNeto * WPlazo1 * WTasa) / 36000
        Call Redondeo(WImpoInteres)
        WNeto = WNeto + WImpoInteres
    End If
    
    WIva1 = 0
    WIva2 = 0
    
    WImpoIb = 0
    WImpoIbTucu = 0
    WImpoIbCiudad = 0
    
    WImpoPorceIb = 0
    WImpoPorceIbTucu = 0
    WImpoPorceIbCiudad = 0
    
    ZFechaCompa = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    If ZFechaCompa >= "20071201" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
            WPorceIb = IIf(IsNull(rstCliente!PorceIb), "0", rstCliente!PorceIb)
            rstCliente.Close
        End If
        
        If ZZIb <> 2 Then
            WImpoIb = WNeto * (WPorceIb / 100)
            Call Redondeo(WImpoIb)
            WImpoPorceIb = WPorceIb
        End If
    
            Else
    
        Select Case WCodIb
            Case 0, 1
                Select Case Val(WCodIva)
                    Case 1
                        WImpoIb = WNeto * 0.025
                        WImpoPorceIb = 2.5
                    Case 2, 4, 5, 6
                        WImpoIb = WNeto * 0.03
                        WImpoPorceIb = 3
                    Case Else
                        WImpoIb = 0
                End Select
                Call Redondeo(WImpoIb)
            Case Else
                WImpoIb = 0
        End Select
        
    End If
    
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WPorceCm05Tucu = IIf(IsNull(rstCliente!PorceCm05Tucu), "0", rstCliente!PorceCm05Tucu)
        rstCliente.Close
    End If
    If WPorceCm05Tucu = 0 Then
        WPorceCm05Tucu = 1
    End If
    Select Case WCodIbTucu
        Case 1, 2, 3
            WImpoIbTucu = WNeto * 0.0175 * WPorceCm05Tucu
            Call Redondeo(WImpoIbTucu)
            WImpoPorceIbTucu = 1.75
        Case 4
            WImpoIbTucu = WNeto * 0.035
            Call Redondeo(WImpoIbTucu)
            WImpoPorceIbTucu = 3.5
        Case 5
            WImpoIbTucu = WNeto * 0.025
            Call Redondeo(WImpoIbTucu)
            WImpoPorceIbTucu = 2.5
        Case Else
            WImpoIbTucu = 0
    End Select
    
    If WImpoIbTucu < 10 Then
        WImpoIbTucu = 0
    End If
    
    Select Case WCodIbCiudad
        Case 1
            WImpoIbCiudad = WNeto * 0.035
            Call Redondeo(WImpoIbCiudad)
            WImpoPorceIbCiudad = 3.5
        Case 2
            WImpoIbCiudad = WNeto * 0.06
            Call Redondeo(WImpoIbCiudad)
            WImpoPorceIbCiudad = 6
        Case Else
            WImpoIbCiudad = 0
    End Select
    
    Compara = WNeto
    Call Redondeo(Compara)
    If Compara < 100 Then
        WImpoIb = 0
    End If
    If Compara < 300 Then
       WImpoIbCiudad = 0
    End If
    
    If ReteIb.ListIndex = 2 Then
        WImpoIb = 0
    End If
    
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.21
            WIva2 = WNeto * 0.105
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Case 3, 4, 5
            WIva1 = 0
            WIva2 = 0
        Case Else
            WIva1 = WNeto * 0.21
            Call Redondeo(WIva1)
    End Select
            
    If WImpoIb <> 0 Then
        Call Convierte1_datos(Str$(WImpoIb), Auxi)
        ImpoIb.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIb.Caption = "0.00"
    End If
    
    If WImpoIbTucu <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbTucu), Auxi)
        ImpoIbTucu.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbTucu.Caption = "0.00"
    End If
    
    If WImpoIbCiudad <> 0 Then
        Call Convierte1_datos(Str$(WImpoIbCiudad), Auxi)
        ImpoIbCiudad.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIbCiudad.Caption = "0.00"
    End If
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Dto.Caption = "0.00"
    End If
    
    If WImpoInteres <> 0 Then
        Call Convierte1_datos(Str$(WImpoInteres), Auxi)
        Interes.Caption = Pusing("###,###.##", Auxi)
            Else
        Interes.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2 + WImpoIb + WImpoIbTucu + WImpoIbCiudad
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    PrgDevolPesos.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    
    ClaveCtacte = "02" + Auxi + "01"
    spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        rstCtacte.Close
        m$ = "Comprobante ya emitido"
        a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
        Exit Sub
    End If

    ClaveCtacte = "05" + Auxi + "01"
    spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        rstCtacte.Close
        m$ = "Comprobante ya emitido"
        a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
        Exit Sub
    End If
        
    If Trim(Cae.Text) <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Cae = " + "'" + Cae.Text + "'"
        spCtacte = ZSql
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            rstCtacte.Close
            m$ = "Nro de Cae ya utilizado"
            a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
            Exit Sub
        End If
    End If

    If ReteIb.ListIndex = 0 Then
        m$ = "Se debe informar si se debe retener o no Ingresos Brutos"
        aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
        Exit Sub
    End If

    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
            
            WRenglon = WRenglon + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 7
            Articulo = UCase(DBGrid1.Text)
            WTipoProDy = Left$(Articulo, 2)
                
            DBGrid1.Col = 2
            Entrada = Val(DBGrid1.Text)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            DBGrid1.Col = 5
            Tipopro = DBGrid1.Text
            
            If Cantidad <> 0 Or Entrada <> 0 Then
                If Tipopro = "" Then
                    m$ = "No se ha informado el tipo de producto"
                    aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    Renglon = 1
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                    Exit Sub
                End If
            End If
                    
        Next iRow
        
    Next a
    
    If Trim(Cae.Text) = "" Then
        ZZGrabaFactura = ""
        Call Calcula_Cae
        If ZZGrabaFactura <> "S" Then
            Exit Sub
        End If
    End If

    Cliente.Text = UCase(Cliente.Text)

    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
            
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1

    DBGrid1.Col = 0
    DBGrid1.Text = ""

    Call Calcula_Click
    
    If Trim(Cae.Text) = "" Then
        ZZGrabaFactura = ""
        Call Calcula_Cae
        If ZZGrabaFactura <> "S" Then
            Exit Sub
        End If
    End If
    
    Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
    Rem     WImporte = WNeto
    Rem    WNeto = WNeto / 1.21
    Rem    Call Redondeo(WNeto)
    Rem    WIva1 = WImporte - WNeto
    Rem    WIva2 = 0
    Rem End If
    
    WTipo = "02"
    WNumero = Numero.Text
    WRenglon = "01"
    WCliente = Cliente.Text
    WFecha = Fecha.Text
    WEstado = "0"
    Rem Wvencimiento = Wvencimiento
    Rem WVencimiento1 = WVencimiento1
    Call Convierte_datos(Str$(Total), Auxi)
    XTotal = Str$(WTotal * -1)
    XTotalUs = Str$((WTotal * -1) / Val(Paridad.Text))
    XSaldo = Str$(WTotal * -1)
    XSaldoUs = Str$((WTotal * -1) / Val(Paridad.Text))
    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
    WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
    WImpre = "DV"
    XNet = Str$(WNeto * -1)
    XIva1 = Str$(WIva1 * -1)
    XIva2 = Str$(WIva2 * -1)
    XImpoIb = Str$(WImpoIb * -1)
    XImpoIbTucu = Str$(WImpoIbTucu * -1)
    XImpoIbCiudad = Str$(WImpoIbCiudad * -1)
    XSeguro = ""
    XFlete = ""
    WPedido = ""
    WRemito = Remito.Text
    WOrden = ""
    WParidad = Paridad.Text
    WProvincia = WProvincia
    XVendedor = Str$(WVendedor)
    XRubro = Str$(WRubro)
    WComprobante = ""
    WAceptada = ""
    WCosto = ""
    WImporte1 = ""
    WImporte2 = ""
    WImporte3 = ""
    WImporte4 = ""
    WImporte5 = ""
    WImporte6 = ""
    WImporte7 = ""
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = "02" + Auxi + "01"
    XEmpresa = "1"
    WDate = Date$
    
    XParam = "'" + WClave + "','" _
                + WTipo + "','" + WNumero + "','" _
                + WRenglon + "','" + WCliente + "','" _
                + WFecha + "','" + WEstado + "','" _
                + Wvencimiento + "','" + WVencimiento1 + "','" _
                + XTotal + "','" + XTotalUs + "','" _
                + XSaldo + "','" + XSaldoUs + "','" _
                + WOrdFecha + "','" + WOrdVencimiento + "','" _
                + WOrdVencimiento1 + "','" + WImpre + "','" _
                + Wempresa + "','" _
                + XNet + "','" + XIva1 + "','" _
                + XIva2 + "','" + WPedido + "','" _
                + WRemito + "','" + WOrden + "','" _
                + WParidad + "','" + WProvincia + "','" _
                + XVendedor + "','" + XRubro + "','" _
                + WComprobante + "','" + WAceptada + "','" _
                + WCosto + "','" _
                + WImporte1 + "','" + WImporte2 + "','" _
                + WImporte3 + "','" + WImporte4 + "','" _
                + WImporte5 + "','" + WImporte6 + "','" _
                + WImporte7 + "','" + WDate + "','" _
                + XSeguro + "','" + XFlete + "','" _
                + XImpoIb + "'"
    
    spCtacte = "AltaCtacte " + XParam
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZImpreNumero = Str$(Val(WNumero) - 100000)
    Call Ceros(ZZImpreNumero, 8)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql + " Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql + " fechaCae = " + "'" + vtocae.Text + "',"
    ZSql = ZSql + " ImpoIbTucu = " + "'" + XImpoIbTucu + "',"
    ZSql = ZSql + " ImpoIbCiudad = " + "'" + XImpoIbCiudad + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                 
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    Rem GRABA REGISTRO 50 %
    
    WAdicional = 0
    WNumero8 = ""
    WImporte8 = 0
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
        rstCliente.Close
    End If
    
    If WAdicional > 0 Then
        If Val(Wempresa) = 8 Then
            OPEN_FILE_Ctacte8
            OPEN_FILE_Numero8
            OPEN_FILE_Esta8
        End If
    End If
    
    If WAdicional > 0 Then
    
         With rstNumero8
            .Index = "Codigo"
            Claveven$ = "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                WNumero8 = Str$(!Numero + 1)
                    Else
                WNumero8 = "1"
            End If
        End With
        
        With rstNumero8
            .Index = "Codigo"
            .Seek "=", "01"
            If .NoMatch = False Then
                .Edit
                !Numero = Val(WNumero8)
                .Update
            End If
        End With
        
        WImporte8 = (WNeto * WAdicional)
        
        With rstCtacte8
            .Index = "Clave"
            .AddNew
            !Tipo = "02"
            !Numero = WNumero8
            !Renglon = "00"
            !Cliente = Cliente.Text
            !Fecha = Fecha.Text
            !Estado = "0"
            !Vencimiento = "  /  /    "
            !Vencimiento1 = "  /  /    "
            Call Convierte_datos(Str$(Total), Auxi)
            !Total = (WNeto * WAdicional) * -1
            !Totalus = ((WNeto * WAdicional) * -1) / Val(Paridad.Text)
            !Saldo = (WNeto * WAdicional) * -1
            !Saldous = ((WNeto * WAdicional) * -1) / Val(Paridad.Text)
            !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            !OrdVencimiento = "00000000"
            !OrdVencimiento1 = "00000000"
            !Impre = "DV"
            !Neto = (WNeto * WAdicional) * -1
            !Iva1 = 0
            !Iva2 = 0
            !Pedido = 0
            !Remito = 0
            !Orden = ""
            !Paridad = Val(Paridad.Text)
            !Provincia = WProvincia
            !vendedor = WVendedor
            !Rubro = WRubro
            !Comprobante = ""
            !Aceptada = ""
            !Costo = 0
            !Importe1 = 0
            !Importe2 = 0
            !Importe3 = 0
            !Importe4 = 0
            !Importe5 = 0
            !Importe6 = 0
            !Importe7 = 0
            Auxi = WNumero8
            Call Ceros(Auxi, 8)
            !Clave = "02" + Auxi + "00"
            !WDate = Date$
            !TipoDescarga = 1
            .Update
        End With
        
    End If
    
    ZAdicional = Str$(WAdicional)
    ZAdicional = Pusing("######", ZAdicional)
    
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = "02" + Auxi + "01"
    Sql1 = "UPDATE Ctacte SET "
    Sql2 = "Adicional = " + "'" + ZAdicional + "',"
    Sql3 = "Numero8 = " + "'" + WNumero8 + "'"
    Sql4 = " Where Clave = " + "'" + WClave + "'"
    spCtacte = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    If WAdicional > 0 Then
        Auxi = WNumero8
        Call Ceros(Auxi, 8)
        WClave = "02" + Auxi + "00"
        With rstCtacte8
            .Index = "Clave"
            .Seek "=", WClave
            If .NoMatch = True Then
                m$ = "No se ha podido generar correctamente la factura complementaria"
                aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    Else
                If Abs(WImporte8) <> Abs(!Total) Or Cliente.Text <> !Cliente Then
                    m$ = "No se ha podido generar correctamente la factura complementaria"
                    aa% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                End If
            End If
        End With
    End If
    
    
    
    
    
    
    
    
    
    Erase VectorCosto
                    
    Renglon = 0
    WRenglon = 0
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0
    
    DBGrid1.Refresh
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
            
            WRenglon = WRenglon + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
            
            DBGrid1.Col = 7
            Articulo = UCase(DBGrid1.Text)
            WTipoProDy = Left$(Articulo, 2)
            
            DBGrid1.Col = 3
            Precio = Val(DBGrid1.Text)
                
            DBGrid1.Col = 2
            Entrada = Val(DBGrid1.Text)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
            
            DBGrid1.Col = 5
            Tipopro = DBGrid1.Text
            
            DBGrid1.Col = 6
            PartiOri = DBGrid1.Text
            Lote = DBGrid1.Text
            
            If WTipoProDy <> "PT" Then
            
                WArtiDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                
                WEntra = "N"
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArtiDy + "'"
                ZSql = ZSql + " and Laudo.PartiOri = " + "'" + PartiOri + "'"
                ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    With rstLaudo
                        .MoveFirst
                        Lote = Str$(rstLaudo!Laudo)
                        WEntra = "S"
                        rstLaudo.Close
                    End With
                End If
                    
                If WEntra = "N" Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArtiDy + "'"
                    ZSql = ZSql + " and Guia.PartiOri = " + "'" + PartiOri + "'"
                    ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        With rstMovguia
                            .MoveFirst
                            Lote = Str$(rstMovguia!Lote)
                            WEntra = "S"
                            rstMovguia.Close
                        End With
                    End If
                    
                End If
                
            End If
                
            If Cantidad <> 0 Or Entrada <> 0 Then
            
                WArti = Tipopro + Mid$(Articulo, 3, 10)
                If WTipoProDy = "PT" Then
                
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WLinea = rstTerminado!Linea
                        rstTerminado.Close
                    End If
                    
                        Else
                        
                    If WTipoProDy = "DY" Then
                        WLinea = 16
                            Else
                        If WTipoProDy = "DS" Then
                            WLinea = 16
                                Else
                            If WTipoProDy = "DQ" Then
                                WLinea = 22
                                    Else
                                WLinea = 5
                            End If
                        End If
                    End If
                    
                End If
             
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                WTipo = "02"
                WNumero = Numero.Text
                XRenglon = Str$(Renglon)
                WArticulo = Tipopro + Mid$(Articulo, 3, 10)
                XCantidad = Str$(Cantidad)
                XPrecio = Str$(Precio)
                XPrecioUs = Str$(Precio / Val(Paridad.Text))
                XImporte = Str$(Precio * Cantidad * -1)
                XImporteUs = Str$((Precio * Cantidad * -1) / Val(Paridad.Text))
                WCliente = Cliente.Text
                WParidad = Paridad.Text
                XVendedor = Str$(WVendedor)
                XRubro = Str$(WRubro)
                XLinea = Str$(WLinea)
                XCosto2 = WCosto1
                XCosto1 = WCosto
                WCoeficiente = ""
                WPedido = ""
                WFecha = Fecha.Text
                WImporte1 = ""
                WImporte2 = ""
                WImporte3 = ""
                WImporte4 = ""
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XArticulo = Tipopro + Mid$(Articulo, 3, 6)
                WRemito = ""
                WClave = "02" + Auxi1 + Auxi
                WDate = Date$
                XCanti = ""
                XImpo = ""
                XImpoUs = ""
                
                XMarca = ""
                If Val(Wempresa) = 1 Then
                    Select Case Planta.ListIndex
                        Case 1
                            XMarca = "X"
                        Case 2
                            XMarca = "X"
                        Case 3
                            XMarca = "X"
                        Case Else
                            XMarca = ""
                    End Select
                End If
                
                WLote1 = Lote
                WCanti1 = Str$(Cantidad)
                WLote2 = "0"
                WCanti2 = "0"
                Wlote3 = "0"
                WCanti3 = "0"
                WLote4 = "0"
                WCanti4 = "0"
                WLote5 = "0"
                WCanti5 = "0"
                WEntrada = Str$(Entrada)
                WTipopro = Tipopro
                WHoja = ""
                If WTipoProDy <> "PT" Then
                    XTipoproDy = "M"
                    XArticuloDy = Tipopro + "-" + Right$(Articulo, 7)
                        Else
                    XTipoproDy = "T"
                    XArticuloDy = "  -   -   "
                End If
                
                XParam = "'" + WClave + "','" _
                    + WTipo + "','" + WNumero + "','" _
                    + XRenglon + "','" + WArticulo + "','" _
                    + XCantidad + "','" + XPrecio + "','" _
                    + XPrecioUs + "','" + XImporte + "','" _
                    + XImporteUs + "','" + WCliente + "','" _
                    + WParidad + "','" + XVendedor + "','" _
                    + XRubro + "','" + XLinea + "','" _
                    + XCosto1 + "','" + XCosto2 + "','" _
                    + WCoeficiente + "','" + WPedido + "','" _
                    + WFecha + "','" + WImporte1 + "','" _
                    + WImporte2 + "','" + WImporte3 + "','" _
                    + WImporte4 + "','" + WOrdFecha + "','" _
                    + XArticulo + "','" + WRemito + "','" _
                    + WDate + "','" + XCanti + "','" _
                    + XImpo + "','" + XImpoUs + "','" _
                    + XMarca + "','" _
                    + WLote1 + "','" + WCanti1 + "','" _
                    + WLote2 + "','" + WCanti2 + "','" _
                    + Wlote3 + "','" + WCanti3 + "','" _
                    + WLote4 + "','" + WCanti4 + "','" _
                    + WLote5 + "','" + WCanti5 + "','" _
                    + WEntrada + "','" _
                    + WTipopro + "','" + WHoja + "','" _
                    + XTipoproDy + "','" + XArticuloDy + "'"
                
                spEstadistica = "AltaEstadisticaDev " + XParam
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                VectorCosto(Renglon, 1) = WArticulo
                VectorCosto(Renglon, 2) = WClave
                
                If Val(Wempresa) = 1 Then
                    If XMarca = "X" Then
                        
                        XEmpresa = Wempresa
                        If Val(Wempresa) = 1 Then
                            Select Case Planta.ListIndex
                                Case 1
                                    Wempresa = "0011"
                                    txtOdbc = "Empresa11"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case 2
                                    Wempresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case 3
                                    Wempresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                            End Select
                        End If
                        
                        XMarca = ""
                        XParam = "'" + WClave + "','" _
                            + WTipo + "','" + WNumero + "','" _
                            + XRenglon + "','" + WArticulo + "','" _
                            + XCantidad + "','" + XPrecio + "','" _
                            + XPrecioUs + "','" + XImporte + "','" _
                            + XImporteUs + "','" + WCliente + "','" _
                            + WParidad + "','" + XVendedor + "','" _
                            + XRubro + "','" + XLinea + "','" _
                            + XCosto1 + "','" + XCosto2 + "','" _
                            + WCoeficiente + "','" + WPedido + "','" _
                            + WFecha + "','" + WImporte1 + "','" _
                            + WImporte2 + "','" + WImporte3 + "','" _
                            + WImporte4 + "','" + WOrdFecha + "','" _
                            + XArticulo + "','" + WRemito + "','" _
                            + WDate + "','" + XCanti + "','" _
                            + XImpo + "','" + XImpoUs + "','" _
                            + XMarca + "','" _
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + Wlote3 + "','" + WCanti3 + "','" _
                            + WLote4 + "','" + WCanti4 + "','" _
                            + WLote5 + "','" + WCanti5 + "','" _
                            + WEntrada + "','" _
                            + WTipopro + "','" + WHoja + "','" _
                            + XTipoproDy + "','" + XArticuloDy + "'"
                    
                        spEstadistica = "AltaEstadisticaDev " + XParam
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                                
                        Call Conecta_Empresa
                        
                    End If
                    
                End If
                
                If WAdicional > 0 Then
                    Auxi1 = Str$(WNumero8)
                    Call Ceros(Auxi1, 8)
                    With rstEsta8
                        .Index = "Clave"
                        .AddNew
                        !Tipo = "02"
                        !Numero = WNumero8
                        !Renglon = Renglon
                        !Articulo = Articulo
                        !Cantidad = Cantidad * -1
                        !Precio = Precio * WAdicional
                        !PrecioUs = (Precio * WAdicional) / Val(Paridad.Text)
                        !Importe = Precio * Cantidad * WAdicional * -1
                        !ImporteUs = (Precio * Cantidad * WAdicional * -1) / Val(Paridad.Text)
                        !Cliente = Cliente.Text
                        !Paridad = Val(Paridad.Text)
                        !vendedor = WVendedor
                        !Rubro = WRubro
                        !Linea = WLinea
                        !Costo1 = 0
                        !Costo2 = 0
                        !Coeficiente = 0
                        !Pedido = 0
                        !Fecha = Fecha.Text
                        !Importe1 = 0
                        !Importe2 = 0
                        !Importe3 = 0
                        !Importe4 = 0
                        !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        !WArticulo = Left$(Articulo, 8)
                        !Remito = ""
                        !Clave = "02" + Auxi1 + Auxi
                        !WDate = Date$
                        !TipoDescarga = 1
                        !lote1 = 0
                        !lote2 = 0
                        !lote3 = 0
                        !lote4 = 0
                        !lote5 = 0
                        !Canti1 = 0
                        !Canti2 = 0
                        !Canti3 = 0
                        !Canti4 = 0
                        !Canti5 = 0
                        .Update
                    End With
                End If
                
                Auxiliar(Renglon, 1) = Articulo
                Auxiliar(Renglon, 2) = Cantidad
                Auxiliar(Renglon, 3) = Precio
                Auxiliar(Renglon, 4) = WRenglon
                Auxiliar(Renglon, 5) = Lote
                Auxiliar(Renglon, 6) = Tipopro
                Auxiliar(Renglon, 7) = XTipoproDy
                Auxiliar(Renglon, 8) = XArticuloDy
                Auxiliar(Renglon, 9) = PartiOri
                
                Dife = Entrada - Cantidad
                
                If Dife > 0 Then
                
                    WArti = "PT-99999-999"
                    spTerminado = "ConsultaTerminado " + "'" + WArti + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WLinea = rstTerminado!Linea
                        rstTerminado.Close
                            Else
                        WLinea = 50
                    End If
                    
                    ZZZNumero = Str$(Val(Numero.Text) + 200000)
                
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                
                    Auxi1 = ZZZNumero
                    Call Ceros(Auxi1, 8)
                    WTipo = "02"
                    WNumero = ZZZNumero
                    XRenglon = Str$(Renglon)
                    WArticulo = WArti
                    XCantidad = Str$(Dife)
                    XPrecio = Str$(Precio)
                    XPrecioUs = Str$(Precio / Val(Paridad.Text))
                    XImporte = Str$(Precio * Dife * -1)
                    XImporteUs = Str$((Precio * Dife * -1) / Val(Paridad.Text))
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = ""
                    XCosto1 = ""
                    WCoeficiente = ""
                    WPedido = ""
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = "PT-99999"
                    WRemito = ""
                    WClave = "02" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    XMarca = ""
                    WLote1 = ""
                    WCanti1 = ""
                    WLote2 = "0"
                    WCanti2 = "0"
                    Wlote3 = "0"
                    WCanti3 = "0"
                    WLote4 = "0"
                    WCanti4 = "0"
                    WLote5 = "0"
                    WCanti5 = "0"
                    WEntrada = ""
                    WTipopro = ""
                    WHoja = ""
                    XTipoproDy = "T"
                    XArticuloDy = "  -   -   "
                
                    XParam = "'" + WClave + "','" _
                        + WTipo + "','" + WNumero + "','" _
                        + XRenglon + "','" + WArticulo + "','" _
                        + XCantidad + "','" + XPrecio + "','" _
                        + XPrecioUs + "','" + XImporte + "','" _
                        + XImporteUs + "','" + WCliente + "','" _
                        + WParidad + "','" + XVendedor + "','" _
                        + XRubro + "','" + XLinea + "','" _
                        + XCosto1 + "','" + XCosto2 + "','" _
                        + WCoeficiente + "','" + WPedido + "','" _
                        + WFecha + "','" + WImporte1 + "','" _
                        + WImporte2 + "','" + WImporte3 + "','" _
                        + WImporte4 + "','" + WOrdFecha + "','" _
                        + XArticulo + "','" + WRemito + "','" _
                        + WDate + "','" + XCanti + "','" _
                        + XImpo + "','" + XImpoUs + "','" _
                        + XMarca + "','" _
                        + WLote1 + "','" + WCanti1 + "','" _
                        + WLote2 + "','" + WCanti2 + "','" _
                        + Wlote3 + "','" + WCanti3 + "','" _
                        + WLote4 + "','" + WCanti4 + "','" _
                        + WLote5 + "','" + WCanti5 + "','" _
                        + WEntrada + "','" _
                        + WTipopro + "','" + WHoja + "','" _
                        + XTipoproDy + "','" + XArticuloDy + "'"
                    
                    spEstadistica = "AltaEstadisticaDev " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                    
            End If
                                    
        Next iRow
        
    Next a
    
    For DA = 1 To Renglon
    
        If Val(Auxiliar(DA, 2)) <> 0 Then
    
            Articulo = Auxiliar(DA, 1)
            Cantidad = Auxiliar(DA, 2)
            Precio = Auxiliar(DA, 3)
            WRenglon = Auxiliar(DA, 4)
            Lote = Auxiliar(DA, 5)
            Tipopro = Auxiliar(DA, 6)
            XTipoproDy = Auxiliar(DA, 7)
            XArticuloDy = Auxiliar(DA, 8)
            PartiOri = Auxiliar(DA, 9)
            
            Select Case Tipopro
                Case "NK", "RE"
                    Rem XCodigo = Tipopro + Mid$(Articulo, 3, 10)
                    Rem spTerminado = "ConsultaTerminado " + "'" + XCodigo + "'"
                    Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    Rem If rstTerminado.RecordCount > 0 Then
                    Rem     WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    Rem     WCodigo = XCodigo
                    Rem     WEntradas = Str$(rstTerminado!Entradas + Cantidad)
                    Rem     WLinea = rstTerminado!Linea
                    Rem     WDate = Date$
                    Rem     rstTerminado.Close
                    Rem
                    Rem     XParam = "'" + XCodigo + "','" _
                    Rem                 + WEntradas + "','" _
                    Rem             + WDate + "'"
                    Rem
                    Rem     spTerminado = "ModificaTerminadoEntradas " + XParam
                    Rem     Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    Rem
                    Rem     If WControla = 0 Then
                    Rem         XParam = "'" + Lote + "','" _
                    Rem                     + XCodigo + "'"
                    Rem         spHoja = "ListaHojaProducto " + XParam
                    Rem         Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    Rem         If rstHoja.RecordCount > 0 Then
                    Rem             WClave = rstHoja!Clave
                    Rem             WSaldo = Str$(rstHoja!Saldo + Cantidad)
                    Rem             WDate = Date$
                    Rem             WMarca = ""
                    Rem             rstHoja.Close
                    Rem
                    Rem             XParam = "'" + WClave + "','" _
                    Rem                         + WDate + "','" _
                    Rem                         + WSaldo + "','" _
                    Rem                         + WMarca + "'"
                    Rem             Rem spHoja = "ModificaHojaSaldo2 " + XParam
                    Rem             Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    Rem         End If
                    Rem     End If
                    Rem End If
                    
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Then
                        Select Case Planta.ListIndex
                            Case 1
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                    End If
                
                    WArticuloNk = "NK" + Mid$(Articulo, 3, 10)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE EntDev SET "
                    ZSql = ZSql + "NroDev = " + "'" + Numero.Text + "',"
                    ZSql = ZSql + "Laboratorio = Laboratorio + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + "Saldo = Saldo - " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Remito.Text + "'"
                    ZSql = ZSql + " and Terminado = " + "'" + WArticuloNk + "'"
                    spEntdev = ZSql
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call Conecta_Empresa
                
                Case "DY", "DS", "DQ"
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Then
                        Select Case Planta.ListIndex
                            Case 1
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                    End If
                    
                    If Tipopro = "DS" Then
                        WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                        ArticuloDk = "NS" + Right$(Articulo, 10)
                        WArtiDk = "NS-" + Right$(Articulo, 7)
                        XArticuloDy = WArti
                            Else
                        If Tipopro = "DY" Then
                            WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                            ArticuloDk = "DK" + Right$(Articulo, 10)
                            WArtiDk = "DK-" + Right$(Articulo, 7)
                            XArticuloDy = WArti
                                Else
                            WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                            ArticuloDk = "NQ" + Right$(Articulo, 10)
                            WArtiDk = "NQ-" + Right$(Articulo, 7)
                            XArticuloDy = WArti
                        End If
                    End If
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE EntDev SET "
                    ZSql = ZSql + "Laboratorio = Laboratorio + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + "Saldo = Saldo - " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Remito.Text + "'"
                    ZSql = ZSql + " and Terminado = " + "'" + ArticuloDk + "'"
                    spEntdev = ZSql
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArtiDk + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = rstArticulo!Codigo
                        WEntradas = Str$(rstArticulo!Entradas)
                        WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                        XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                        spArticulo = "ModificaArticuloMovimientos " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = rstArticulo!Codigo
                        WEntradas = Str$(rstArticulo!Entradas + Val(Cantidad))
                        WSalidas = Str$(rstArticulo!Salidas)
                        XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                        spArticulo = "ModificaArticuloMovimientos " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                    If WControla = 0 And Val(Lote) <> 0 Then
                        XParam = "'" + Lote + "','" _
                                    + XArticuloDy + "'"
                        spLaudo = "ListaLaudoArticulo " + XParam
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            WClave = rstLaudo!Clave
                            WSaldo = Str$(rstLaudo!Saldo + Cantidad)
                            WDate = Date$
                            rstLaudo.Close
                            
                            XParam = "'" + WClave + "','" _
                                    + WDate + "','" _
                                    + WSaldo + "'"
                            spLaudo = "ModificaLaudoSaldo " + XParam
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                
                                    Else
                                    
                            XParam = "'" + XArticuloDy + "','" _
                                        + Lote + "'"
                            spMovguia = "ListaMovguiaLote " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WClave = rstMovguia!Clave
                                WSaldo = Str$(rstMovguia!Saldo + Cantidad)
                                WDate = Date$
                                rstMovguia.Close
                                
                                XParam = "'" + WClave + "','" _
                                        + WDate + "','" _
                                        + WSaldo + "'"
                                spMovguia = "ModificaMovguiaSaldo " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                Case "DK", "NS", "NQ"
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Then
                        Select Case Planta.ListIndex
                            Case 1
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                    End If
                    
                    If Tipopro = "DS" Then
                        WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                        ArticuloDk = "DS" + Right$(Articulo, 10)
                        WArtiDk = "DS-" + Right$(Articulo, 7)
                        XArticuloDy = WArti
                            Else
                        WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                        ArticuloDk = "DK" + Right$(Articulo, 10)
                        WArtiDk = "DK-" + Right$(Articulo, 7)
                        XArticuloDy = WArti
                    End If
                    
                    Sql1 = "UPDATE EntDev SET "
                    Sql2 = " Saldo = 0"
                    Sql3 = " Where EntDev.Terminado = " + "'" + ArticuloDk + "'"
                    Sql4 = " and EntDev.PartiOri = " + "'" + PartiOri + "'"
                    spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    
                    WLaudo = "995000"
                    spLaudo = "ListaLaudoDevol"
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveLast
                            WLaudo = Str$(rstLaudo!Laudo + 1)
                        End With
                        rstLaudo.Close
                            Else
                        WLaudo = "995000"
                    End If
                    
                    WPartida = WLaudo
                    WCantidad = Cantidad
        
                    WRenglon = "1"
                    WFecha = Fecha.Text
                    WOrden = ""
                    WLiberada = Str$(WCantidad)
                    WDevuelta = "0"
                    WLote = WLaudo
                    WRechazo = ""
                    WActualiza = "N"
                    WMarca = ""
                    WInforme = ""
                    WSaldo = Str$(WCantidad)
                    WOrigenOri = ""
                    WPartiOri = WPartida
                    WEnvase = ""
                    
                    Auxi1 = Str$(WLaudo)
                    Call Ceros(Auxi1, 6)
                    Auxi2 = Str$(WRenglon)
                    Call Ceros(Auxi2, 2)
                
                    WClave = Auxi1 + Auxi2
                    WDate = Date$
        
                    XParam = "'" + WClave + "','" _
                                 + WLaudo + "','" _
                                 + WRenglon + "','" _
                                 + WFecha + "','" _
                                 + WArtiDk + "','" _
                                 + WLiberada + "','" _
                                 + WDevuelta + "','" _
                                 + WOrden + "','" _
                                 + WMarca + "','" _
                                 + WLote + "','" _
                                 + WRechazo + "','" _
                                 + WInforme + "','" _
                                 + WActualiza + "','" _
                                 + WDate + "','" _
                                 + WSaldo + "','" _
                                 + WOrigenOri + "','" _
                                 + WPartiOri + "','" _
                                 + WEnvase + "'"
                    
                    Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
                    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    XParam = "'" + WLaudo + "','" _
                                 + WFechaord + "'"
                     
                    Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call Conecta_Empresa
                
                Case Else
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Then
                        Select Case Planta.ListIndex
                            Case 1
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                    End If
                
                    spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        WCodigo = Articulo
                        WEntradas = Str$(rstTerminado!Entradas + Cantidad)
                        WLinea = rstTerminado!Linea
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                                     + WEntradas + "','" _
                                     + WDate + "'"
                                                    
                        spTerminado = "ModificaTerminadoEntradas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            
                        If WControla = 0 Then
                            XParam = "'" + Lote + "','" _
                                        + Articulo + "'"
                            spHoja = "ListaHojaProducto " + XParam
                            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                            If rstHoja.RecordCount > 0 Then
                                WClave = rstHoja!Clave
                                WSaldo = Str$(rstHoja!Saldo + Cantidad)
                                WDate = Date$
                                WMarca = ""
                                rstHoja.Close
                                
                                XParam = "'" + WClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "','" _
                                            + WMarca + "'"
                                spHoja = "ModificaHojaSaldo2 " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    
                                            Else
                                        
                                XParam = "'" + Articulo + "','" _
                                                + Lote + "'"
                                spMovguia = "ListaMovguiaLote1 " + XParam
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    WClave = rstMovguia!Clave
                                    WSaldo = Str$(rstMovguia!Saldo + Cantidad)
                                    WDate = Date$
                                    rstMovguia.Close
                                    
                                    XParam = "'" + WClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                    spMovguia = "ModificaMovguiaSaldo " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                                    
                            End If
                        End If
                        
                    End If
                    
                    spTerminado = "ConsultaTerminado " + "'" + "NK" + Mid$(Articulo, 3, 10) + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = "NK" + Mid$(Articulo, 3, 10)
                        WSalidas = Str$(rstTerminado!Salidas + Cantidad)
                        WLinea = rstTerminado!Linea
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                                                                           
                        spTerminado = "ModificaTerminadoSalidas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                        
                    WArticuloNk = "NK" + Mid$(Articulo, 3, 10)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE EntDev SET "
                    ZSql = ZSql + "NroDev = " + "'" + Numero.Text + "',"
                    ZSql = ZSql + "Laboratorio = Laboratorio + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + "Saldo = Saldo - " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Remito.Text + "'"
                    ZSql = ZSql + " and Terminado = " + "'" + WArticuloNk + "'"
                    spEntdev = ZSql
                    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call Conecta_Empresa
                    
            End Select
            
            XEmpresa = Wempresa
            If Val(Wempresa) = 1 Then
                Select Case Planta.ListIndex
                    Case 1
                        Wempresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        Wempresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 3
                        Wempresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
            End If
        
            ZSql = ""
            ZSql = ZSql + "UPDATE LiberaTerminado SET "
            ZSql = ZSql + "ImpreVentas = " + "'" + "S" + "'"
            ZSql = ZSql + " Where PedidoDevol = " + "'" + Remito.Text + "'"
            ZSql = ZSql + " and Producto = " + "'" + Articulo + "'"
            ZSql = ZSql + " and Cliente = " + "'" + Cliente.Text + "'"
            spLiberaTerminado = ZSql
            Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
            Call Conecta_Empresa
        
        End If
        
    Next DA
    
    spNumero = "ConsultaNumero " + "'" + "53" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        WCodigo = "53"
        WNumero = Numero.Text
        XParam = "'" + WCodigo + "','" _
                 + WNumero + "'"
        spNumero = "ModificaNumero " + XParam
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Rem dada
    For Ciclo = 1 To 100
    
        If VectorCosto(Ciclo, 1) <> "" Then
        
            ZZZProducto = VectorCosto(Ciclo, 1)
            ZZClave = VectorCosto(Ciclo, 2)
            
            ZZZCosto = 0
            Call Calcula_CostoFactura(ZZZProducto, ZZZCosto)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Estadistica SET "
            ZSql = ZSql + " Costo1 = " + "'" + Str$(ZZZCosto) + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    
    
    Rem Call Impresion
    Call ImpresionFe
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Numero.SetFocus
        
End Sub


Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    WPrecio.Caption = ""
    
    WArticulo.SetFocus
    
End Sub


Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cae.Text = ""
    vtocae.Text = "00/00/0000"
    Remito.Text = ""
    ReteIb.ListIndex = 0
    
    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WLote.Text = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
  
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 7
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    
    spNumero = "ConsultaNumero " + "'" + "53" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
            Else
        Numero.Text = "1"
    End If
    
    Graba.Enabled = True
    Borra.Enabled = True
    Ingresa.Enabled = True
    
    Numero.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WArticulo.Text = UCase(WArticulo.Text)
        
        WCliente = Cliente.Text
        WTerminado = WArticulo.Text
        WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
        WClave = Cliente.Text + WArticulo.Text
        WClaveMp = Cliente.Text + WArti
        
        If Left$(WArticulo.Text, 2) <> "PT" Then
            XTipoPro = "M"
                Else
            XTipoPro = "T"
        End If
        
        Select Case XTipoPro
            Case "M"
                spPreciosMp = "ConsultaPreciosMp " + "'" + WClaveMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    WEntra = "S"
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                    rstPreciosMp.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
                XArticulo = Left$(WArticulo.Text, 3) + Right$(WArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + XArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WEntra = "S"
                    WDescripcion.Caption = rstPrecios!Descripcion
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                    rstPrecios.Close
                    WCantidad.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
        End Select
        
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM LiberaTerminado"
        ZSql = ZSql & " Where LiberaTerminado.Producto = " + "'" + WArticulo.Text + "'"
        ZSql = ZSql & " and LiberaTerminado.PedidoDevol = " + "'" + Remito.Text + "'"
        ZSql = ZSql & " and LiberaTerminado.Cliente = " + "'" + Cliente.Text + "'"
        spLiberaTerminado = ZSql
        Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstLiberaTerminado.RecordCount > 0 Then
            WTipopro.Text = rstLiberaTerminado!Tipo
            If XTipoPro = "M" Then
                WLote.Text = rstLiberaTerminado!PartiOri
                    Else
                WLote.Text = rstLiberaTerminado!Partida
            End If
            rstLiberaTerminado.Close
        End If
        
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WEntrada.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEntrada_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WEntrada.Text = Pusing("###,###.##", WEntrada.Text)
        WTipopro.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipopro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Left$(WArticulo.Text, 2) = "DY" Then
            If WTipopro.Text = "DY" Or WTipopro.Text = "DK" Then
                WLote.SetFocus
            End If
        End If
        If Left$(WArticulo.Text, 2) = "DS" Then
            If WTipopro.Text = "DS" Or WTipopro.Text = "NS" Then
               WLote.SetFocus
            End If
        End If
        If Left$(WArticulo.Text, 2) = "DQ" Then
            If WTipopro.Text = "DQ" Or WTipopro.Text = "NQ" Then
                WLote.SetFocus
            End If
        End If
        
        If Left$(WArticulo.Text, 2) <> "DY" And Left$(WArticulo.Text, 2) <> "DS" And Left$(WArticulo.Text, 2) <> "DQ" Then
             If WTipopro.Text = "PT" Or WTipopro.Text = "NK" Or WTipopro.Text = "RE" Then
                WLote.SetFocus
            End If
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WLote_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Select Case WTipopro.Text
            Case "PT"
                WEntra = "N"
            
                XEmpresa = Wempresa
                If Val(Wempresa) = 1 Then
                    Select Case Planta.ListIndex
                        Case 1
                            Wempresa = "0011"
                            txtOdbc = "Empresa11"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            Wempresa = "0007"
                            txtOdbc = "Empresa07"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 3
                            Wempresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                End If
                
                WArti = "NK-" + Right$(WArticulo.Text, 9)
                WEstado = ""
                Sql1 = "Select *"
                Sql2 = " FROM EntDev"
                Sql3 = " Where EntDev.Terminado = " + "'" + WArti + "'"
                Sql4 = " and EntDev.Lote = " + "'" + WLote.Text + "'"
                spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                    WEntra = "S"
                    WEstado = IIf(IsNull(rstEntdev!Estado), "", rstEntdev!Estado)
                    WSaldo = rstEntdev!Saldo
                    rstEntdev.Close
                End If
                
                If WEntra = "N" Then
                    Call Conecta_Empresa
                    m$ = "El Articulo o Lote no coincide con la Entrada de Devolucion"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                    Exit Sub
                End If
                
                If WSaldo <= 0 Then
                    Call Conecta_Empresa
                    m$ = "El Articulo no posee la cantidad de producto liberado para emitir el comprobante"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                    Exit Sub
                End If
                
                If WEstado <> "PT" Then
                    Call Conecta_Empresa
                    m$ = "El tipo de producto no puede ser PT"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                    Exit Sub
                End If
            
                WEntra = "N"
            
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
            
                If WControla = 0 Then
                    XParam = "'" + WLote.Text + "','" _
                                 + WArticulo.Text + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WEntra = "S"
                        rstHoja.Close
                    End If
                    
                    If WEntra = "N" Then
                        XParam = "'" + WArticulo.Text + "','" _
                                + WLote.Text + "'"
                        spMovguia = "ListaMovguiaLote1 " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            WEntra = "S"
                            rstMovguia.Close
                        End If
                    End If
                    
                        Else
                        
                    WEntra = "S"
                    
                End If
                
                Call Conecta_Empresa
        
                If WEntra = "N" Then
                    m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
                
            Case "DY", "DK", "DS", "NS", "DQ", "NQ"
                If WTipopro.Text = "DY" Or WTipopro.Text = "DK" Then
                    WArticulo.Text = UCase(WArticulo.Text)
                    WArti = "DK-" + Right$(WArticulo.Text, 9)
                    WEntra = "N"
                        Else
                    If WTipopro.Text = "DS" Or WTipopro.Text = "NS" Then
                        WArticulo.Text = UCase(WArticulo.Text)
                        WArti = "NS-" + Right$(WArticulo.Text, 9)
                        WEntra = "N"
                            Else
                        If WTipopro.Text = "DQ" Or WTipopro.Text = "NQ" Then
                            WArticulo.Text = UCase(WArticulo.Text)
                            WArti = "NQ-" + Right$(WArticulo.Text, 9)
                            WEntra = "N"
                        End If
                    End If
                End If
                
                XEmpresa = Wempresa
                If Val(Wempresa) = 1 Then
                    Select Case Planta.ListIndex
                        Case 1
                            Wempresa = "0011"
                            txtOdbc = "Empresa11"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            Wempresa = "0007"
                            txtOdbc = "Empresa07"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 3
                            Wempresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                End If
                
                Sql1 = "Select *"
                Sql2 = " FROM EntDev"
                Sql3 = " Where EntDev.Terminado = " + "'" + WArti + "'"
                Sql4 = " and EntDev.PartiOri = " + "'" + WLote.Text + "'"
                spEntdev = Sql1 + Sql2 + Sql3 + Sql4
                Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
                If rstEntdev.RecordCount > 0 Then
                    WEntra = "S"
                    rstEntdev.Close
                End If
                    
                Call Conecta_Empresa
                    
                If WEntra = "N" Then
                    m$ = WArticulo.Text + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
                
                Call Conecta_Empresa
            
            Case Else
                WArti = WTipopro.Text + Mid$(WArticulo.Text, 3, 10)
                WEntra = "S"
                
                WControla = 0
                spTerminado = "ConsultaTerminado " + "'" + WArticulo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                    rstTerminado.Close
                End If
                
                
                If WControla = 0 Then
                
                    WEntra = "N"
                    
                    XEmpresa = Wempresa
                    If Val(Wempresa) = 1 Then
                        Select Case Planta.ListIndex
                            Case 1
                                Wempresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 2
                                Wempresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case 3
                                Wempresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                        End Select
                    End If
                
                    XParam = "'" + WLote.Text + "','" _
                                 + WArti + "'"
                    spHoja = "ListaHojaProducto " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHoja.RecordCount > 0 Then
                        WEntra = "S"
                        rstHoja.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                End If
            
                If WEntra = "N" Then
                    m$ = WArti + " Producto inexistente o Lote nro. " + WLote.Text + " inexistente"
                    G% = MsgBox(m$, 0, "Nota de Credito por Devolucion")
                        Else
                    Call Alta_Vector
                    Call Ingresa_Click
                    Call Calcula_Click
                    WArticulo.SetFocus
                End If
                
        End Select
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spClientes = "ConsultaCliente " + "'" + Claveven$ + "'"
            Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
            If rstClientes.RecordCount > 0 Then
                Cliente.Text = rstClientes!Cliente
                DesCliente.Caption = rstClientes!Razon
                WAdicional = IIf(IsNull(rstClientes!Adicional), "0", rstClientes!Adicional)
                WPago1 = rstClientes!Pago1
                WPago2 = rstClientes!Pago2
                WVendedor = rstClientes!vendedor
                WProvincia = rstClientes!Provincia
                WRubro = rstClientes!Rubro
                WCodIva = rstClientes!Iva
                WCodIb = IIf(IsNull(rstClientes!Ib), "0", rstClientes!Ib)
                WCodIbTucu = IIf(IsNull(rstClientes!IbTucu), "0", rstClientes!IbTucu)
                Rem WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstClientes!Razon
                WDireccion = rstClientes!Direccion
                WLocalidad = rstClientes!Localidad
                WPostal = rstClientes!Postal
                WCuit = rstClientes!Cuit
                WDirentrega = rstClientes!DirEntrega
                rstClientes.Close
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
            End If
            Ayuda.Visible = False
            
        Case 1
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            If Mid$(Claveven$, 7, 2) = "DY" Or Mid$(Claveven$, 7, 2) = "DS" Or Mid$(Claveven$, 7, 2) = "DQ" Then
            
                spPreciosMp = "ConsultaPreciosMp " + "'" + Claveven$ + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    DBGrid1.Col = 0
                    DBGrid1.Text = Mid$(Claveven$, 7, 3) + "00" + Right$(Claveven$, 7)
                    DBGrid1.Col = 7
                    DBGrid1.Text = Mid$(Claveven$, 7, 3) + "00" + Right$(Claveven$, 7)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                    WArticulo.Text = Mid$(Claveven$, 7, 3) + "00" + Right$(Claveven$, 7)
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                    rstPreciosMp.Close
                End If
                
                XArticulo = Left$(WArticulo, 3) + Right$(WArticulo, 7)
                spArticulo = "ConsultaArticulo " + "'" + XArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    WDescripcion.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                    
                Call Alta_Vector
                WLinea.Text = WAnterior + 1
                If Val(WLinea.Text) > 0 Then
                    DBGrid1.Row = Val(WLinea.Text) - 1
                End If
                    
                Call DBGrid1.SetFocus
                WCantidad.SetFocus
            
                    Else
            
                spPrecios = "ConsultaPrecios " + "'" + Claveven$ + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstPrecios!Terminado
                    DBGrid1.Col = 7
                    DBGrid1.Text = rstPrecios!Terminado
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                
                    WArticulo.Text = rstPrecios!Terminado
                    WDescripcion.Caption = rstPrecios!Descripcion
                    WPrecio.Caption = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                    
                    rstPrecios.Close
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WCantidad.SetFocus
                    
                End If
                
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6, 7
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

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Consumidor Final"
    Iva(4) = "Exento"
    Iva(5) = "Monotributo"
    Iva(6) = "No Catalogado"
    
    Planta.Clear
    
    Planta.AddItem "Planta I (CO/PG)"
    Planta.AddItem "Planta VII (FA)"
    Planta.AddItem "Planta V (PT/BI)"
    Planta.AddItem "Planta II (40000)"
    
    Planta.ListIndex = 0
    
    ReteIb.Clear
    
    ReteIb.AddItem ""
    ReteIb.AddItem "Calcula"
    ReteIb.AddItem "No Calcula"
    
    ReteIb.ListIndex = 0
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 7, 0 To 40)

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
For i = 0 To 7
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Entrada a Pta."
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Tipo"
             DBGrid1.Columns(newcnt).Width = 700
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "Lote"
             DBGrid1.Columns(newcnt).Width = 1100
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 7
             DBGrid1.Columns(newcnt).Caption = ""
             DBGrid1.Columns(newcnt).Width = 10
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i
 
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cae.Text = ""
    vtocae.Text = "00/00/0000"
    
    WLinea.Text = ""
    WArticulo.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WPrecio.Caption = ""
    WEntrada.Text = ""
    WTipopro.Text = ""
    WLote.Text = ""
    Remito.Text = ""
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "53" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
            Else
        Numero.Text = "1"
    End If
    
    Rem DBGrid1.FirstRow = 0
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
            spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
    
    Numero.SetFocus
    
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
            DBGrid1.Text = Pusing("###,###.##", WEntrada.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WTipopro.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
                
            DBGrid1.Col = 7
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Row = Renglon
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
            DBGrid1.Text = Pusing("###,###.##", WEntrada.Text)
            
            DBGrid1.Col = 5
            DBGrid1.Text = WTipopro.Text
            
            DBGrid1.Col = 6
            DBGrid1.Text = WLote.Text
                
            DBGrid1.Col = 7
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 7
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    Erase Auxiliar
    
    XParam = "'" + "02" + "','" _
                + Numero.Text + "'"
    
    spEstadistica = "ConsultaEstadistica1 " + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = !Articulo
                    Auxi1 = !Articulo
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Abs(!Entrada)
               Rem************BY NAN*************
               Rem  10-10-2012 MODIFICADO PARA PESOS
               Rem   CAMBIA !pRECIOUS POR PRECIO
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(!Precio))
               Rem BY NAN
                    DBGrid1.Col = 4
                    DBGrid1.Text = Abs(!Cantidad)
                    
                    DBGrid1.Col = 5
                    DBGrid1.Text = !Tipopro
                    
                    DBGrid1.Col = 6
                    DBGrid1.Text = Pusing("######", Str$(!lote1))
                
                    DBGrid1.Col = 7
                    DBGrid1.Text = !Articulo
                    Auxi1 = !Articulo
                
                    Rem Paridad.Text = !Lote
                    
                    Auxiliar(Renglon, 1) = Auxi1
    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    XRenglon = Renglon
    Renglon = 0
    
    For DA = 1 To XRenglon
    
        Auxi1 = Auxiliar(DA, 1)
        
        If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DS" Or Left$(Auxi1, 2) = "DQ" Then
        
            WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
            spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Renglon = Renglon + 1
                    
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                    
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                    
                DBGrid1.Col = 1
                DBGrid1.Text = rstArticulo!Descripcion
                rstArticulo.Close
            End If
        
                Else
                
            ClavePrecios = Cliente.Text + "PT" + Mid$(Auxi1, 3, 10)
        
            spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                    
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                    
                DBGrid1.Col = 1
                DBGrid1.Text = rstPrecios!Descripcion
                rstPrecios.Close
            End If
            
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
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 10) * 10
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = False
    Borra.Enabled = False
    Ingresa.Enabled = False
    
    Call Calcula_Click

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "02" + Auxi + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Paridad.Text = rstCtacte!Paridad
            Paridad.Text = Pusing("#,###.###", Paridad.Text)
            Remito.Text = rstCtacte!Remito
            
            
            Cae.Text = IIf(IsNull(rstCtacte!Cae), "", rstCtacte!Cae)
            vtocae.Text = IIf(IsNull(rstCtacte!fechaCae), "  /  /    ", rstCtacte!fechaCae)
            
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!vendedor
                WProvincia = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WCodIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
            End If
            Call Proceso_Click
                Else
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Devoluciones")
            Rem     Numero.SetFocus
            Rem         Else
            Rem     Graba.Enabled = True
            Rem     Borra.Enabled = True
            Rem     Ingresa.Enabled = True
            Rem     WNumero = Numero.Text
            Rem     Numero.Text = WNumero
            Rem     Cliente.SetFocus
            Rem End If
            Graba.Enabled = True
            Borra.Enabled = True
            Ingresa.Enabled = True
            WNumero = Numero.Text
            Numero.Text = WNumero
            Planta.SetFocus
                
        End If
    End If
End Sub

Private Sub Planta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Remito.SetFocus
    End If
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = Wempresa
        If Val(Wempresa) = 1 Then
            Select Case Planta.ListIndex
                Case 1
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        End If
    
        spEntdev = "ListaEntdev " + "'" + Remito.Text + "'"
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        If rstEntdev.RecordCount > 0 Then
            Cliente.Text = rstEntdev!Cliente
            Cliente.Text = UCase(Cliente.Text)
            rstEntdev.Close
            
            Renglon = 0
            
            spEntdev = "ListaEntdev " + "'" + Remito.Text + "'"
            Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
            If rstEntdev.RecordCount > 0 Then
                With rstEntdev
                    .MoveFirst
                    Do
                        If .EOF = False Then
                
                            Renglon = Renglon + 1
            
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 7
                            Select Case Left$(rstEntdev!Terminado, 2)
                                Case "NK"
                                    DBGrid1.Text = "PT" + Mid$(rstEntdev!Terminado, 3, 10)
                                    Auxi1 = "PT" + Mid$(rstEntdev!Terminado, 3, 10)
                                Case "DK"
                                    DBGrid1.Text = "DY" + Mid$(rstEntdev!Terminado, 3, 10)
                                    Auxi1 = "DY" + Mid$(rstEntdev!Terminado, 3, 10)
                                Case "NS"
                                    DBGrid1.Text = "DS" + Mid$(rstEntdev!Terminado, 3, 10)
                                    Auxi1 = "DS" + Mid$(rstEntdev!Terminado, 3, 10)
                                Case "NQ"
                                    DBGrid1.Text = "DQ" + Mid$(rstEntdev!Terminado, 3, 10)
                                    Auxi1 = "DQ" + Mid$(rstEntdev!Terminado, 3, 10)
                                Case Else
                                    DBGrid1.Text = "PT" + Mid$(rstEntdev!Terminado, 3, 10)
                                    Auxi1 = "PT" + Mid$(rstEntdev!Terminado, 3, 10)
                            End Select
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", rstEntdev!Cantidad)
                
                            If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Then
                                DBGrid1.Col = 6
                                DBGrid1.Text = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                                    Else
                                DBGrid1.Col = 6
                                DBGrid1.Text = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                            End If
                    
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###.##", rstEntdev!Cantidad)
                            
                            Auxiliar(Renglon, 1) = Auxi1
                
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEntdev.Close
            End If
    
            WRenglon = Renglon
            Renglon = 0

            For DA = 1 To WRenglon
    
                Auxi1 = Auxiliar(DA, 1)
    
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 10) * 10
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
    
                If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DK" Then
                    WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                            Else
                    spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = rstTerminado!Descripcion
                        rstTerminado.Close
                    End If
                End If
        
            Next DA
                    
            Rem DBGrid1.FirstRow = 0
    
            Renglon = Renglon + 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
    
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
    
            Renglon = Renglon + 1
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
    
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
    
            Renglon = Renglon - 2
            Lugar1 = Int((Renglon - 1) / 10) * 10
            Lugar2 = Renglon - Lugar1
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
    
            Call Conecta_Empresa
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!vendedor
                WProvincia = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WCodIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                rstCliente.Close
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
            End If
            Fecha.SetFocus
            
                Else
                
            Call Conecta_Empresa
            Remito.SetFocus
            
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Cliente.Text = rstCliente!Cliente
            DesCliente.Caption = rstCliente!Razon
            WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
            WPago1 = rstCliente!Pago1
            WPago2 = rstCliente!Pago2
            WVendedor = rstCliente!vendedor
            WProvincia = rstCliente!Provincia
            WRubro = rstCliente!Rubro
            WCodIva = rstCliente!Iva
            WCodIb = IIf(IsNull(rstCliente!Ib), "0", rstCliente!Ib)
            WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
            WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            WDirentrega = rstCliente!DirEntrega
            rstCliente.Close
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            Fecha.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
            End If
            Call Calcula_FechaVto
            Vencimiento.Text = Wvencimiento
            Paridad.SetFocus
                Else
            m$ = "Formato de fecha invalida"
            a% = MsgBox(m$, 0, "Emision de Devolucion")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Paridad.Text) = 0 Then
            m$ = "No exsite paridad cargada para esta fecha"
            a% = MsgBox(m$, 0, "Emision de Devolucion")
            Paridad.SetFocus
                Else
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
        End If
    End If
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Remito.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Impresion()

    Close #55
        
    If Val(Wempresa) = 1 Then
        Rem Open "dadaNC.txt" For Output As #55
        Open "lpt1" For Output As #55
            Else
        Rem Open "dada.txt" For Output As #55
        Open "lpt1" For Output As #55
        Print #55, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #55, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
    End If
    
    Rem Width #55, 255

    Print #55, Chr$(27) + Chr$(40) + "19U";
    Print #55, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #55, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    ImporteIb = Val(ImpoIb.Caption)
    ImporteIbTucu = Val(ImpoIbTucu.Caption)
    ImporteIbCiudad = Val(ImpoIbCiudad.Caption)
    ImpoNeto = Val(Neto.Caption)
    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption))
    Impotot = Val(Total.Caption)

    For XX% = 1 To 2
    
        If Val(Wempresa) = 1 Then
            Rem Print #55, ""
        End If

        Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, Tab(55); "NOTA DE CREDITO"
        Print #55, ""
        Print #55, Tab(59); Fecha.Text
        Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, Tab(8); WRazon
        Print #55, Tab(8); WDireccion
        Print #55, Tab(8); Left$(WLocalidad, 33);
        Print #55, Tab(55); Cliente.Text
        Print #55, Tab(8); Provincia(Val(WProvincia)); " ("; WPostal; ")"
        Print #55, ""
        Print #55, Tab(8); Iva(Val(WCodIva));
        Print #55, Tab(48); WCuit
        Print #55, ""
        Print #55, ""
        Print #55, Tab(5); WPago
        Print #55, ""
        Print #55, ""
        Print #55, ""

        Impre = 0

        For a = 0 To 3
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 7
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
            
                DBGrid1.Col = 2
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                    Print #55, Tab(1); Alinea("#####.##", Str$(Cantidad));
                    Print #55, " Kg";
                    Print #55, Tab(15); Left$(Descri, 38);
                    parcial = Precio * Cantidad
                    Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                    Rem    Print #55, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem    Print #55, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem            Else
                    Rem    Precio = Str$(Val(Precio) * 1.21)
                    Rem    Parcial = Str$(Val(Parcial) * 1.21)
                    Rem    Print #55, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem    Print #55, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem End If
                    Print #55, Tab(56); "U$S"; Alinea("####.##", Str$(Precio));
                    Print #55, Tab(68); Alinea("###,###.##", Str$(parcial))
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next a

        Rem ojo
        Rem For aa = Impre To 22
        
        For aa = Impre To 21
                Print #55, ""
        Next aa

        Rem M# = Total# / 100
        Rem GoSub 4630
        
        If Val(Wempresa) = 1 Then
        
            Print #55, Tab(1); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
            Print #55, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
            Print #55, Tab(1); "TIPO DE CAMBIO:";
            Print #55, Alinea("##.##", Paridad.Text);
            Print #55, " I.V.A.:";
            Print #55, Alinea("#,###.##", Str$(ImpoIva));
            Print #55, " TOTAL:";
            Print #55, Alinea("###,###.##", Str$(Impotot))
            Print #55, Tab(1); "NETO:";
            Print #55, Alinea("###,###.##", Str$(ImpoNeto));
            If ImporteIb <> 0 Then
                Print #55, " I.BRUTOS BS.AS.:";
                Print #55, Alinea("#,###.##", Str$(ImporteIb));
            End If
            If ImporteIbTucu <> 0 Then
                Print #55, " I.BRUTOS TUCUMAN:";
                Print #55, Alinea("#,###.##", Str$(ImporteIbTucu));
            End If
            If ImporteIbCiudad <> 0 Then
                Print #55, " I.BRUTOS CIUDAD:";
                Print #55, Alinea("#,###.##", Str$(ImporteIbCiudad))
                    Else
                Print #55, ""
            End If
            
                Else
            
            Print #55, Tab(3); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
            Print #55, Tab(3); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
            Print #55, Tab(3); "TIPO DE CAMBIO:";
            Print #55, Alinea("##.##", Paridad.Text);
            Print #55, " I.V.A.:";
            Print #55, Alinea("#,###.##", Str$(ImpoIva));
            Print #55, " TOTAL:";
            Print #55, Alinea("###,###.##", Str$(Impotot))
            Print #55, Tab(3); "IMPORTE NETO:";
            Print #55, Alinea("###,###.##", Str$(ImpoNeto));
            If ImporteIb <> 0 Then
                Print #55, " PERCEPCION DE INGRESOS BRUTOS:";
                Print #55, Alinea("#,###.##", Str$(ImporteIb))
                    Else
                Print #55, ""
            End If
            
        End If
        
        Rem Print #55, ""
        Rem Print #55, ""
        Print #55, ""
        Print #55, ""
        Print #55, ""
        If Val(Wempresa) <> 1 Then
            Print #55, ""
            Print #55, ""
        End If
        
        Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(XNeto))

        If Val(Interes.Caption) <> 0 Then
                Print #55, Tab(56); "Interes";
                Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(Val(Interes.Caption)))
                                                  Else
                Print #55, ""
        End If

        If Val(Dto.Caption) <> 0 Then
                Print #55, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(Val(Dto.Caption)))
                        Else
                Print #55, ""
        End If

        Print #55, Tab(3); M1;
        Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(Val(Neto.Caption)))
        Print #55, Tab(3); M2;
        If Val(Iva1.Caption) <> 0 Then
                Print #55, Tab(61); "21";
                Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(Val(Iva1.Caption)))
                        Else
                Print #55, ""
        End If

        Select Case XX
                Case 1
                        Print #55, Tab(3); "ORIGINAL";
                Case 2
                        Print #55, Tab(3); "DUPLICADO";
                Case 3
                        Print #55, Tab(3); "TRIPLICADO";
                Case Else
        End Select
        
        If Val(ImpoIbCiudad.Caption) <> 0 Then
               Print #55, Tab(14); "P.Ciudad";
               Print #55, Tab(65); "U$S"; Alinea("##,###.##", ImpoIbCiudad.Caption);
        End If
        If Val(ImpoIbTucu.Caption) <> 0 Then
                Print #55, Tab(36); "P.Tucuman";
                Print #55, Tab(46); "U$S"; Alinea("##,###.##", ImpoIbTucu.Caption);
        End If
        If Val(ImpoIb.Caption) <> 0 Then
                Print #55, Tab(60); "I.B.";
                Print #55, Tab(65); "U$S"; Alinea("##,###.##", Str$(Val(ImpoIb.Caption)))
                        Else
                If Val(Iva2.Caption) <> 0 Then
                    Print #55, Tab(60); "10.5";
                    Print #55, Tab(65); "U$S"; Alinea("##,###.##", Str$(Val(Iva2.Caption)))
                        Else
                    Print #55, ""
                End If
        End If

        Print #55, Tab(65); "U$S"; Alinea("###,###.##", Str$(Val(Total.Caption)));
        Print #55, Chr$(12)

    Next XX%

    Close #55
    DA = 0

End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCliente = "ListaClienteConsulta"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
            
                    DA = Len(rstCliente!Razon) - WEspacios
                
                    For aa = 1 To DA
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = rstCliente!Cliente
                            IngresaItem = Auxi + "    " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
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
        rstCliente.Close
    End If
    End If

End Sub





Private Sub Calcula_Cae()

    Dim WSAA As Object, WSFEv1 As Object
    
    On Error GoTo ManejoError
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    ' Crear objeto interface Web Service Autenticaci?n y Autorizaci?n
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    'Debug.Print WSAA.InstallDir
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
        
    ZPath = ""
    Select Case Val(Wempresa)
        Case 1
            ZNombre = "c:\salva\surfactan"
            ZCuit = "30549165083"
        Case Else
            ZNombre = "c:\salva\pellital"
            ZCuit = "30610524598"
    End Select
    
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Rem Certificado = "..\..\reingart.crt" ' certificado de prueba
    Rem ClavePrivada = "..\..\reingart.key" ' clave privada de prueba
    
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    proxy = "" '"usuario:clave@localhost:8000"
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n

    ' Imprimir el ticket de acceso, ToKen y Sign de autorizaci?n
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este per?odo se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electr?nica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    'Debug.Print WSFEv1.InstallDir
    
    ' Setear tocken y sing de autorizaci?n (pasos previos)
    WSFEv1.Token = WSAA.Token
    WSFEv1.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturaci?n
    proxy = "" ' "usuario:clave@localhost:8000"
    wsdl = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
    cache = ""    'Rem Path
        
    ok = WSFEv1.Conectar(cache, wsdl, proxy, "") ' homologaci?n
    Debug.Print WSFEv1.Version
    
    ' mostrar bit?cora de depuraci?n:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = 3
    Select Case Val(Wempresa)
        Case 1
            punto_vta = 9
        Case Else
            punto_vta = 6
    End Select
    
    Cbte_Nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    
    If Cbte_Nro = "" Then
        Cbte_Nro = 0                ' no hay comprobantes emitidos
            Else
        Cbte_Nro = CLng(Cbte_Nro)   ' convertir a entero largo
    End If
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WCuit = rstCliente!Cuit
        Call Eval
        rstCliente.Close
    End If
    
    Rem 1-PRODUCTO    2-SERVICIOS     3-PRODUCTOS Y SERVICIOS
    Concepto = 1
    
    Rem TIPO DE DOCUMENTO
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    
    Rem NUMERO DE DOCUMENTO
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    If Val(Numero.Text) - 100000 <> Cbte_Nro + 1 Then
        m$ = "El numero de comprobante no es igua al correlativo indicado por la afip " + Str$(LastCBTE + 1)
        a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
        Exit Sub
    End If
    
    WImpoIb = 0
    WImpoIbTucu = 0
    WImpoIbCiudad = 0
    
    ZZImpoIb = Val(ImpoIb.Caption)
    ZZImpoIbTucu = Val(ImpoIbTucu.Caption)
    ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption)
    ZZImpoNeto = Val(Neto.Caption)
    ZZImpoIva = Val(Iva1.Caption) + Val(Iva2.Caption)
    ZZImpoTotal = Val(Total.Caption)
    
    Rem NUMERO DE DOCUMENTO
    Cbte_Nro = Cbte_Nro + 1
    cbt_desde = Cbte_Nro
    cbt_hasta = Cbte_Nro
    
    Rem IMPORTE TOTAL
    IMP_TOTAL = ZZImpoTotal
    
    Rem IMPORTE DE CONCEPTOS NO GRAVADOS POR EL IVA
    imp_tot_conc = 0
    
    Rem IMPORTE NETO
    imp_neto = ZZImpoNeto
    
    Rem IMPORTE IVA
    imp_iva = ZZImpoIva
    
    Rem suma de importes de otros impuestos
    imp_trib = ZZImpoIb + ZZImpoIbTucu + ZZImpoIbCiudad
    
    Rem IMPORTE EXENTO DE IVA
    imp_op_ex = 0
    
    Rem FECHA
    ZZFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    fecha_cbte = ZZFecha
    
    Rem VENCIMIENTO
    Rem ZZFecha = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
    Rem fecha_venc_pago = ZZFecha
    fecha_venc_pago = ""
    
    Rem FECHAS DE SERVICIOS PARA SERVICIOS
    ' Fechas del per?odo del servicio facturado (solo si concepto = 1?)
    fecha_serv_desde = ""
    fecha_serv_hasta = ""
    
    Rem MONEDA
    moneda_id = "PES"
    
    Rem COTIZACION
    moneda_ctz = 1

    ok = WSFEv1.CrearFactura(Concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, IMP_TOTAL, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz)
    
    ' Agrego los comprobantes asociados:
    Rem If False Then ' solo nc/nd
    Rem     tipo = 19
    Rem     pto_vta = 2
    Rem     nro = 1234
    Rem     ok = WSFEv1.AgregarCmpAsoc(tipo, pto_vta, nro)
    Rem End If
        
    ' Agrego impuestos varios
    If ZZImpoIb <> 0 Then
        id = 2
        Desc = "Percepcion I.Brutos Bs.As."
        base_imp = ZZImpoNeto
        alic = WImpoPorceIb
        Importe = ZZImpoIb
        ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, Importe)
    End If
    
    If ZZImpoIbCiudad <> 0 Then
        id = 2
        Desc = "Percepcion I.Brutos CABA"
        base_imp = ZZImpoNeto
        alic = WImpoPorceIbCiudad
        Importe = ZZImpoIbCiudad
        ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, Importe)
    End If
    
    If ZZImpoIbTucu <> 0 Then
        id = 2
        Desc = "Percepcion I.Brutos Tucuman"
        base_imp = ZZImpoNeto
        alic = WImpoPorceIbTucu
        Importe = ZZImpoIbTucu
        ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, Importe)
    End If

    ' Agrego tasas de IVA
    id = 5 ' 21%
    base_imp = ZZImpoNeto
    Importe = ZZImpoIva
    ok = WSFEv1.AgregarIva(id, base_imp, Importe)
    
    ' Habilito reprocesamiento autom?tico (predeterminado):
    WSFEv1.Reprocesar = True

    ' Solicito CAE:
    Cae = WSFEv1.CAESolicitar()
    
    Debug.Print "Resultado", WSFEv1.Resultado
    Debug.Print "CAE", WSFEv1.Cae

    Debug.Print "Numero de comprobante:", WSFEv1.CbteNro
    
    ' Imprimo pedido y respuesta XML para depuraci?n (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.Cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo

    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & Cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    ' Muestro los errores
    If WSFEv1.ErrMsg <> "" Then
        MsgBox WSFEv1.ErrMsg, vbExclamation, "Error"
    End If
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEv1.Eventos:
        MsgBox evento, vbInformation, "Evento"
    Next
    
    ' Buscar la factura
    cae2 = WSFEv1.CompConsultar(tipo_cbte, punto_vta, Cbte_Nro)
    
    Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.Vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal
    Debug.Print "Resultado:", WSFEv1.Resultado
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & Cae & " vs " & cae2
    Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
        
    If WSFEv1.Resultado = "A" Then
        ZZGrabaFactura = "S"
        Cae.Text = Cae
        If Len(Trim(WSFEv1.Vencimiento)) = 8 Then
            vtocae.Text = Right$(WSFEv1.Vencimiento, 2) + "/" + Mid$(WSFEv1.Vencimiento, 5, 2) + "/" + Left$(WSFEv1.Vencimiento, 4)
                Else
            vtocae.Text = WSFEv1.Vencimiento
        End If
    End If

    Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print WSFEv1.Excepcion
    Debug.Print Err.Description            ' descripci?n error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Print WSFEv1.XmlRequest
            Debug.Print WSFEv1.XmlResponse
            Debug.Print WSFEv1.traceback
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEv1.XmlRequest
    Debug.Assert False
    Debug.Print WSFEv1.traceback

End Sub


Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub

Private Sub Calcula_Barra()
    
    Dim ZZCara(1000) As String
    
    ZZNumero = ""
    Select Case Val(Wempresa)
        Case 1
            ZZNumero = "30549165083"
        Case Else
            ZZNumero = "30610524598"
    End Select
    
    ZZNumero = ZZNumero + "03"
    
    Select Case Val(Wempresa)
        Case 1
            ZZNumero = ZZNumero + "0009"
        Case Else
            ZZNumero = ZZNumero + "0006"
    End Select
    
    ZZNumero = ZZNumero + Trim(Cae.Text)
    
    ZZFechaCae = vtocae.Text
    ZZOrdFechaCae = Right$(ZZFechaCae, 4) + Mid$(ZZFechaCae, 4, 2) + Left$(ZZFechaCae, 2)
    ZZNumero = ZZNumero + ZZOrdFechaCae
    
    ZZCara(0) = "!"
    ZZCara(1) = Chr$(34)
    ZZCara(2) = "#"
    ZZCara(3) = "$"
    ZZCara(4) = "%"
    ZZCara(5) = "&"
    ZZCara(6) = "?"
    ZZCara(7) = "("
    ZZCara(8) = ")"
    ZZCara(9) = "*"
    ZZCara(10) = "+"
    ZZCara(11) = ","
    ZZCara(12) = "-"
    ZZCara(13) = "."
    ZZCara(14) = "/"
    ZZCara(15) = "0"
    ZZCara(16) = "1"
    ZZCara(17) = "2"
    ZZCara(18) = "3"
    ZZCara(19) = "4"
    ZZCara(20) = "5"
    ZZCara(21) = "6"
    ZZCara(22) = "7"
    ZZCara(23) = "8"
    ZZCara(24) = "9"
    ZZCara(25) = ":"
    ZZCara(26) = ";"
    ZZCara(27) = "<"
    ZZCara(28) = "="
    ZZCara(29) = ">"
    ZZCara(30) = "?"
    ZZCara(31) = "@"
    ZZCara(32) = "A"
    ZZCara(33) = "B"
    ZZCara(34) = "C"
    ZZCara(35) = "D"
    ZZCara(36) = "E"
    ZZCara(37) = "F"
    ZZCara(38) = "G"
    ZZCara(39) = "H"
    ZZCara(40) = "I"
    ZZCara(41) = "J"
    ZZCara(42) = "K"
    ZZCara(43) = "L"
    ZZCara(44) = "M"
    ZZCara(45) = "N"
    ZZCara(46) = "O"
    ZZCara(47) = "P"
    ZZCara(48) = "Q"
    ZZCara(49) = "R"
    ZZCara(50) = "S"
    ZZCara(51) = "T"
    ZZCara(52) = "U"
    ZZCara(53) = "V"
    ZZCara(54) = "W"
    ZZCara(55) = "X"
    ZZCara(56) = "Y"
    ZZCara(57) = "Z"
    ZZCara(58) = "["
    ZZCara(59) = "\"
    ZZCara(60) = "]"
    ZZCara(61) = "^"
    ZZCara(62) = "_"
    ZZCara(63) = "`"
    ZZCara(64) = "a"
    ZZCara(65) = "b"
    ZZCara(66) = "c"
    ZZCara(67) = "d"
    ZZCara(68) = "e"
    ZZCara(69) = "f"
    ZZCara(70) = "g"
    ZZCara(71) = "h"
    ZZCara(72) = "i"
    ZZCara(73) = "j"
    ZZCara(74) = "k"
    ZZCara(75) = "l"
    ZZCara(76) = "m"
    ZZCara(77) = "n"
    ZZCara(78) = "o"
    ZZCara(79) = "p"
    ZZCara(80) = "q"
    ZZCara(81) = "r"
    ZZCara(82) = "s"
    ZZCara(83) = "t"
    ZZCara(84) = "u"
    ZZCara(85) = "v"
    ZZCara(86) = "w"
    ZZCara(87) = "x"
    ZZCara(88) = "y"
    ZZCara(89) = "z"
    ZZCara(90) = "¡"
    ZZCara(91) = "¢"
    ZZCara(92) = "£"
    ZZCara(93) = "¤"
    ZZCara(94) = "¥"
    ZZCara(95) = "¦"
    ZZCara(96) = "§"
    ZZCara(97) = "¨"
    ZZCara(98) = "©"
    ZZCara(99) = "ª"
    
    Rem ZZNumero = "3070306062119000260321213344273201008198"
    Rem ZZNumero = "000102030405060708091011121314151617181920"
    Rem ZZNumero = "2122232425262728293031323334353637383940"
    Rem ZZNumero = "4142434445464748495051525354555657585960"
    Rem ZZNumero = "6162636465666768697071727374757677787980"
    Rem ZZNumero = "81828384858687888990919293949596979899"
    Rem ZZNumero = "307030606211900026032121334427320100819"
    
    ZZSumaI = 0
    ZZSumaII = 0
    
    For Ciclo = 1 To 39 Step 2
        ZZSumaI = ZZSumaI + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    ZZSumaI = ZZSumaI * 3
    
    For Ciclo = 2 To 39 Step 2
        ZZSumaII = ZZSumaII + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    
    ZZSuma = ZZSumaI + ZZSumaII
    ZZVerifica = ZZSuma
    ZZDigi = 0
    
    Do
    
        ZZVerifi = Int(ZZVerifica / 10) * 10
        
        If ZZVerifi = ZZVerifica Then
            Exit Do
        End If
        
        ZZDigi = ZZDigi + 1
        
        ZZVerifica = ZZSuma + ZZDigi
        
    Loop
    
    ZZNumero = ZZNumero + Trim(Str$(ZZDigi))
    
    lccar = ""
    barralargo = ZZNumero
    
    For lni = 1 To Len(barralargo) Step 2
        ZZLugar = Val(Mid(barralargo, lni, 2))
        lccar = lccar + ZZCara(ZZLugar)
    Next
    
    Rem barralargo = "{" + lccar + "}"
    barralargo = "(" + lccar + ")"
    
    ZZImpreBarraI = barralargo
    ZZImpreBarraII = ZZNumero

End Sub







Sub ImpresionFe()

    Call Calcula_Barra
    
    Auxi1 = Str$(Val(Numero.Text) - 100000)
    Call Ceros(Auxi1, 8)
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpreFactura"
    Rem ZSql = ZSql + " Where Numero = " + "'" + Auxi1 + "'"
    spImpreFactura = ZSql
    Set rstImpreFactura = db.OpenRecordset(spImpreFactura, dbOpenSnapshot, dbSQLPassThrough)

    ImporteIb = Val(ImpoIb.Caption)
    ImporteIbTucu = Val(ImpoIbTucu.Caption)
    ImporteIbCiudad = Val(ImpoIbCiudad.Caption)
    ImpoNeto = Val(Neto.Caption)
    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption))
    Impotot = Val(Total.Caption)

    Erase ZZVector
    ZZLugarII = 0
    Impre = 0

    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
        
            Suma = Suma + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 7
            Producto = DBGrid1.Text
            
            DBGrid1.Col = 1
            Descri = DBGrid1.Text
            
            DBGrid1.Col = 3
            XPrecio = Val(DBGrid1.Text)
        
            DBGrid1.Col = 2
            Rem BY NAN 20-9-2013
            
          
            Cantidad = Val(IIf(DBGrid1.Text = "", 0, DBGrid1.Text))
           
            
          
          Rem Cantidad = DBGrid1.Text
            
            ZZLugarII = ZZLugarII + 1
            
            ZZVector(ZZLugarII, 1) = Str$(Cantidad)
            ZZVector(ZZLugarII, 2) = Descri
            ZZVector(ZZLugarII, 3) = Str$(XPrecio)
            ZZVector(ZZLugarII, 4) = Str$(XPrecio * Cantidad)
                
        Next iRow
        
    Next a
        
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WPago1 = rstCliente!Pago1
        WPago2 = rstCliente!Pago2
        WVendedor = rstCliente!vendedor
        WProv = rstCliente!Provincia
        WRubro = rstCliente!Rubro
        WCodIva = rstCliente!Iva
        WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
        WCodIb = rstCliente!Ib
        WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
        WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
        WRazon = Trim(rstCliente!Razon)
        WDireccion = Trim(rstCliente!Direccion)
        WLocalidad = Trim(rstCliente!Localidad)
        WPostal = Trim(rstCliente!Postal)
        WCuit = Trim(rstCliente!Cuit)
        rstCliente.Close
    End If
        
    ZZImpoIb = Val(ImpoIb.Caption)
    ZZImpoIbTucu = Val(ImpoIbTucu.Caption)
    ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption)
    ZZImpoNeto = Val(Neto.Caption)
    ZZImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption))
    ZZImpoTotal = Val(Total.Caption)

    For aa = Impre To 16
                
        ZZCantidad = Val(ZZVector(aa, 1))
        ZZDescripcion = ZZVector(aa, 2)
        ZZPrecio = Val(ZZVector(aa, 3))
        ZZParcial = Val(ZZVector(aa, 4))
                    
        If Val(Numero.Text) > 100000 Then
            Auxi1 = Str$(Val(Numero.Text) - 100000)
                Else
            Auxi1 = Numero.Text
        End If
        Call Ceros(Auxi1, 8)
        Auxi2 = Str$(aa)
        Call Ceros(Auxi2, 2)
        
                        
        WWClave = Auxi1 + Auxi2
        WWNumero = Auxi1
        WWRenglon = Str(aa)
        WWFecha = Fecha.Text
        WWCliente = Cliente.Text
        WWRazon = WRazon
        WWDireccion = WDireccion
        WWLocalidad = WLocalidad
        WWOrden = ""
        WWProvincia = Provincia(Val(WProv)) + " (" + WPostal + ")"
        WWPostal = WPostal
        WWIva = Iva(Val(WCodIva))
        WWCuit = WCuit
        wwpago = Trim(WPago) + "   " + Vencimiento.Text
        WWVencimiento = Vencimiento.Text
        WWRemito = Remito.Text
        WWCantidad = Str$(ZZCantidad)
        WWDescripcion = ZZDescripcion
        WWPrecio = Str$(ZZPrecio)
        WWParcial = Str$(ZZParcial)
        WWImpreDespachoI = ImpreDespachoI
        WWImpreDespachoII = ImpreDespachoII
        WWParidad = Paridad.Text
        WWImpoIva = Str$(ZZImpoIva)
        WWImpotot = Str$(ZZImpoTotal)
        WWImpoNeto = Str$(ZZImpoNeto)
        WWImporteIb = Str$(ZZImpoIb)
        WWImporteIbTucuman = Str$(ZZImpoIbTucu)
        WWImporteIbCiudad = Str$(ZZImpoIbCiudad)
        WWPorceDescuento = Str$(WDescuento)
        WWDescuento = Dto.Caption
        WWInteres = Interes.Caption
        WWImprePesos1 = ""
        WWImprePesos2 = ""
        WWNeto = Neto.Caption
        WWNetoII = Str$(XNeto)
        WWIva1 = Iva1.Caption
        WWIva2 = Iva2.Caption
        WWIbCiudad = ImpoIbCiudad.Caption
        WWIbTucuman = ImpoIbTucu.Caption
        WWIb = ImpoIb.Caption
        WWTotal = Total.Caption
        WWImpreComprobante = "NOTA DE CREDITO"
        WWCae = Cae.Text
        WWFechaCae = vtocae.Text
        WWImpreBarraI = ZZImpreBarraI
        WWImpreBarraII = ZZImpreBarraII
                        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreFactura ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Pago ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Parcial ,"
        ZSql = ZSql + "ImpreDespachoI ,"
        ZSql = ZSql + "ImpreDespachoII ,"
        ZSql = ZSql + "Paridad ,"
        ZSql = ZSql + "ImpoIva ,"
        ZSql = ZSql + "Impotot ,"
        ZSql = ZSql + "ImpoNeto ,"
        ZSql = ZSql + "ImporteIb ,"
        ZSql = ZSql + "ImporteIbTucuman ,"
        ZSql = ZSql + "ImporteIbCiudad ,"
        ZSql = ZSql + "PorceDescuento ,"
        ZSql = ZSql + "Descuento ,"
        ZSql = ZSql + "Interes ,"
        ZSql = ZSql + "ImprePesos1 ,"
        ZSql = ZSql + "ImprePesos2 ,"
        ZSql = ZSql + "Neto ,"
        ZSql = ZSql + "NetoII ,"
        ZSql = ZSql + "Iva1 ,"
        ZSql = ZSql + "Iva2 ,"
        ZSql = ZSql + "IbCiudad ,"
        ZSql = ZSql + "IbTucuman ,"
        ZSql = ZSql + "Ib ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "FechaCae ,"
        ZSql = ZSql + "ImpreBarraI ,"
        ZSql = ZSql + "ImpreBarraII ,"
        ZSql = ZSql + "ImpreComprobante )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WWClave + "',"
        ZSql = ZSql + "'" + WWNumero + "',"
        ZSql = ZSql + "'" + WWRenglon + "',"
        ZSql = ZSql + "'" + WWFecha + "',"
        ZSql = ZSql + "'" + WWRazon + "',"
        ZSql = ZSql + "'" + WWDireccion + "',"
        ZSql = ZSql + "'" + WWLocalidad + "',"
        ZSql = ZSql + "'" + WWCliente + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + WWProvincia + "',"
        ZSql = ZSql + "'" + WWPostal + "',"
        ZSql = ZSql + "'" + WWIva + "',"
        ZSql = ZSql + "'" + WWCuit + "',"
        ZSql = ZSql + "'" + wwpago + "',"
        ZSql = ZSql + "'" + WWVencimiento + "',"
        ZSql = ZSql + "'" + WWRemito + "',"
        ZSql = ZSql + "'" + WWCantidad + "',"
        ZSql = ZSql + "'" + WWDescripcion + "',"
        ZSql = ZSql + "'" + WWPrecio + "',"
        ZSql = ZSql + "'" + WWParcial + "',"
        ZSql = ZSql + "'" + Left$(WWImpreDespachoI, 100) + "',"
        ZSql = ZSql + "'" + Left$(WWImpreDespachoII, 100) + "',"
        ZSql = ZSql + "'" + WWParidad + "',"
        ZSql = ZSql + "'" + WWImpoIva + "',"
        ZSql = ZSql + "'" + WWImpotot + "',"
        ZSql = ZSql + "'" + WWImpoNeto + "',"
        ZSql = ZSql + "'" + WWImporteIb + "',"
        ZSql = ZSql + "'" + WWImporteIbTucuman + "',"
        ZSql = ZSql + "'" + WWImporteIbCiudad + "',"
        ZSql = ZSql + "'" + WWPorceDescuento + "',"
        ZSql = ZSql + "'" + WWDescuento + "',"
        ZSql = ZSql + "'" + WWInteres + "',"
        ZSql = ZSql + "'" + WWImprePesos1 + "',"
        ZSql = ZSql + "'" + WWImprePesos2 + "',"
        ZSql = ZSql + "'" + WWNeto + "',"
        ZSql = ZSql + "'" + WWNetoII + "',"
        ZSql = ZSql + "'" + WWIva1 + "',"
        ZSql = ZSql + "'" + WWIva2 + "',"
        ZSql = ZSql + "'" + WWIbCiudad + "',"
        ZSql = ZSql + "'" + WWIbTucuman + "',"
        ZSql = ZSql + "'" + WWIb + "',"
        ZSql = ZSql + "'" + WWTotal + "',"
        ZSql = ZSql + "'" + WWCae + "',"
        ZSql = ZSql + "'" + WWFechaCae + "',"
        ZSql = ZSql + "'" + WWImpreBarraI + "',"
        ZSql = ZSql + "'" + WWImpreBarraII + "',"
        ZSql = ZSql + "'" + WWImpreComprobante + "')"
            
        spImpreFactura = ZSql
        Set rstImpreFactura = db.OpenRecordset(spImpreFactura, dbOpenSnapshot, dbSQLPassThrough)
                        
    Next aa
            
    Listado.WindowTitle = "Factura Electronica"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.CopiesToPrinter = 2
    
    If Val(Wempresa) = 1 Then
        Listado.ReportFileName = "ImpreFacturaLocalPesosNuevo.rpt"
            Else
        Listado.ReportFileName = "ImpreFacturaLocalPesosPelli.rpt"
    End If
                
    Rem Listado.ReportFileName = "Impredada.rpt"
                
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT ImpreFactura.Numero, ImpreFactura.Renglon, ImpreFactura.Fecha, ImpreFactura.Razon, " _
                       + "ImpreFactura.Direccion, ImpreFactura.Localidad, ImpreFactura.Cliente, ImpreFactura.Orden, ImpreFactura.Provincia, ImpreFactura.Iva, ImpreFactura.Cuit, ImpreFactura.Pago, ImpreFactura.Remito, ImpreFactura.Cantidad, ImpreFactura.Descripcion, ImpreFactura.Precio, ImpreFactura.Parcial, ImpreFactura.ImpreDespachoI, ImpreFactura.ImpreDespachoII, ImpreFactura.Paridad, ImpreFactura.ImpoIva, ImpreFactura.Impotot, ImpreFactura.ImpoNeto, ImpreFactura.ImporteIb, ImpreFactura.ImporteIbTucuman, ImpreFactura.ImporteIbCiudad, ImpreFactura.PorceDescuento, ImpreFactura.Descuento, ImpreFactura.Interes, ImpreFactura.Neto, ImpreFactura.Iva1, ImpreFactura.IbCiudad, ImpreFactura.IbTucuman, ImpreFactura.Ib, ImpreFactura.Total, ImpreFactura.ImpreComprobante, ImpreFactura.Cae, ImpreFactura.FechaCae, ImpreFactura.ImpreBarraI, ImpreFactura.ImpreBarraII, ImpreFactura.NetoII " _
                       + "From " _
                       + DSQ + ".dbo.ImpreFactura ImpreFactura " _
                       + "Where " _
                       + "ImpreFactura.Numero >= '0' AND " _
                       + "ImpreFactura.Numero <= '999999'"
                       
    Listado.Destination = 1
    Rem Listado.Destination = 0

    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub



