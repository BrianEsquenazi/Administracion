VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompras 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Comprobantes de Proveedores"
   ClientHeight    =   6660
   ClientLeft      =   405
   ClientTop       =   990
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   6660
   ScaleWidth      =   11190
   Begin VB.TextBox Despacho 
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
      Left            =   8760
      MaxLength       =   20
      TabIndex        =   55
      Text            =   " "
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Cai 
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
      Left            =   9360
      MaxLength       =   14
      TabIndex        =   51
      Text            =   " "
      Top             =   480
      Width           =   1695
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
      Left            =   9840
      MaxLength       =   15
      TabIndex        =   50
      Text            =   " "
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Pago 
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
      Left            =   6840
      TabIndex        =   48
      Top             =   1320
      Width           =   1815
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
      Height          =   1620
      Left            =   5280
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Tipo 
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
      Left            =   600
      TabIndex        =   46
      Text            =   " "
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox TipoComp 
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
      Left            =   840
      TabIndex        =   2
      Text            =   " "
      Top             =   480
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
      Height          =   2595
      ItemData        =   "ComprasAnterior.frx":0000
      Left            =   4560
      List            =   "ComprasAnterior.frx":0007
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   4935
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "ComprasAnterior.frx":0015
      TabIndex        =   16
      Top             =   3600
      Width           =   9375
   End
   Begin MSMask.MaskEdBox Vencimiento1 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   1320
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
   Begin VB.TextBox NroInterno 
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
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
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
      Left            =   8520
      TabIndex        =   42
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Exento 
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
      Left            =   6000
      MaxLength       =   15
      TabIndex        =   15
      Text            =   " "
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Iva27 
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
      Left            =   6000
      MaxLength       =   15
      TabIndex        =   13
      Text            =   " "
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Iva21 
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
      Left            =   6000
      MaxLength       =   15
      TabIndex        =   11
      Text            =   " "
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Ib 
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   45
      Text            =   " "
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Iva5 
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
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   12
      Text            =   " "
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Punto 
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
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   4
      Text            =   " "
      Top             =   480
      Width           =   855
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
      Left            =   6120
      MaxLength       =   8
      TabIndex        =   5
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Letra 
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
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame4 
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
      Height          =   615
      Left            =   3840
      TabIndex        =   29
      Top             =   2880
      Width           =   3015
      Begin VB.OptionButton Contado2 
         Caption         =   "En Cta.Cte."
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
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Contado1 
         Caption         =   "Efectivo"
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
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSMask.MaskEdBox Periodo 
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   960
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
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
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
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   960
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
   Begin VB.TextBox Neto 
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
      Height          =   315
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   10
      Text            =   " "
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   3600
      MaxLength       =   11
      TabIndex        =   1
      Text            =   " "
      Top             =   0
      Width           =   1335
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
      Height          =   345
      Left            =   9600
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
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
      Height          =   345
      Left            =   8520
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   18
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      TabIndex        =   17
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox VtoCai 
      Height          =   285
      Left            =   9720
      TabIndex        =   53
      Top             =   840
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
   Begin VB.Label Label20 
      Caption         =   "Despacho"
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
      Left            =   7440
      TabIndex        =   56
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Vto CAI"
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
      Left            =   8280
      TabIndex        =   54
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "C.A.I."
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
      Left            =   8280
      TabIndex        =   52
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label9 
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
      Left            =   8760
      TabIndex        =   49
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Forma de Pago"
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
      TabIndex        =   47
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nro Interno"
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
      TabIndex        =   44
      Top             =   0
      Width           =   1215
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
      Height          =   285
      Left            =   2280
      TabIndex        =   43
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Importe Total"
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
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Importe No Gravado"
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
      Left            =   3840
      TabIndex        =   40
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Importe Iva 27%"
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
      TabIndex        =   39
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Importe Iva 21%"
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
      TabIndex        =   38
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Importe Perc. I.B."
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
      TabIndex        =   37
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Iva R.G. 3337"
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
      TabIndex        =   36
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Punto"
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
      Left            =   3240
      TabIndex        =   35
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Numero"
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
      Left            =   5040
      TabIndex        =   34
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label13 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Letra"
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
      Left            =   2040
      TabIndex        =   32
      Top             =   480
      Width           =   495
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
      Height          =   285
      Left            =   5040
      TabIndex        =   28
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "Fecha de Iva"
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
      Left            =   3960
      TabIndex        =   27
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha de vencimiento"
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
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Importe Neto"
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
      Width           =   2175
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
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha de Emision"
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
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "PrgCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 4 ' Número máximo de campos del conjunto de registros.
Private Dato As String
Private Auxi As String
Private WImpo As Double
Private WProveedor As String
Private SumaDebito As Double
Private SumaCredito As Double
Private Uno As Double
Private Dos As Double
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstImputac As Recordset
Dim spImputac As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim XParam As String
Dim cParam As String

Sub Calcula_total()

    WImpo = 0
    Call Format_datos
    
    Dato = Neto.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva21.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva5.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva27.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Ib.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Exento.Text
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Total.Caption = WImpo
    Total.Caption = Pusing("#,###,###.##", Total.Caption)
    
End Sub

Sub Alinea_datos()
    Tipo.Text = Str$(TipoComp.ListIndex + 1)
    WTipo = Tipo.Text
    Call Ceros(WTipo, 2)
    Tipo.Text = WTipo
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    Punto.Text = WPunto
    WNumero = Numero.Text
    Call Ceros(WNumero, 8)
    Numero.Text = WNumero
    Letra.Text = Left$(Letra.Text, 1)
End Sub

Sub Imprime_Descripcion()
    With RstProveedor
        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
                Else
            DesProveedor.Caption = ""
        End If
    End With
End Sub

Sub Verifica_datos()
    If Val(Neto.Text) = 0 Then
        Neto.Text = "0"
    End If
    If Val(Iva21.Text) = 0 Then
        Iva21.Text = "0"
    End If
    If Val(Iva5.Text) = 0 Then
        Iva5.Text = "0"
    End If
    If Val(Iva27.Text) = 0 Then
        Iva27.Text = "0"
    End If
    If Val(Ib.Text) = 0 Then
        Ib.Text = "0"
    End If
    If Val(Exento.Text) = 0 Then
        Exento.Text = "0"
    End If
    If Val(Total.Caption) = 0 Then
        Total.Caption = "0"
    End If
    If Val(Paridad.Text) = 0 Then
        Paridad.Text = "0"
    End If
End Sub

Sub Format_datos()
    If Val(Paridad.Text) <> 0 Then
        Paridad.Text = Pusing("#,###.####", Paridad.Text)
    End If
    If Val(Neto.Text) <> 0 Then
        Neto.Text = Pusing("#,###,###.##", Neto.Text)
    End If
    If Val(Iva21.Text) <> 0 Then
        Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
    End If
    If Val(Iva5.Text) <> 0 Then
        Iva5.Text = Pusing("#,###,###.##", Iva5.Text)
    End If
    If Val(Iva27.Text) <> 0 Then
        Iva27.Text = Pusing("#,###,###.##", Iva27.Text)
    End If
    If Val(Ib.Text) <> 0 Then
        Ib.Text = Pusing("#,###,###.##", Ib.Text)
    End If
    If Val(Exento.Text) <> 0 Then
        Exento.Text = Pusing("#,###,###.##", Exento.Text)
    End If
    Total.Caption = Pusing("#,###,###.##", Total.Caption)
End Sub

Sub Imprime_Datos()

    spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
            Proveedor.Text = rstIvaComp!Proveedor
            TipoComp.ListIndex = rstIvaComp!Tipo - 1
            Letra.Text = rstIvaComp!Letra
            Punto.Text = rstIvaComp!Punto
            Numero.Text = rstIvaComp!Numero
            Call Alinea_datos
            Fecha.Text = rstIvaComp!Fecha
            Vencimiento.Text = rstIvaComp!Vencimiento
            Vencimiento1.Text = rstIvaComp!Vencimiento1
            Periodo.Text = rstIvaComp!Periodo
            Neto.Text = Abs(rstIvaComp!Neto)
            Iva21.Text = Abs(rstIvaComp!Iva21)
            Iva5.Text = Abs(rstIvaComp!Iva5)
            Iva27.Text = Abs(rstIvaComp!Iva27)
            Ib.Text = Abs(rstIvaComp!Ib)
            Exento.Text = Abs(rstIvaComp!Exento)
            Call Calcula_total
            Contado1.Value = False
            Contado2.Value = False
            Select Case Val(rstIvaComp!Contado)
                Case 1
                    Contado1.Value = True
                Case 2
                    Contado2.Value = True
                Case Else
            End Select
            Paridad.Text = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
            Pago.ListIndex = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
            Cai.Text = IIf(IsNull(rstIvaComp!Cai), "", rstIvaComp!Cai)
            Cai.Text = Trim(Cai.Text)
            VtoCai.Text = IIf(IsNull(rstIvaComp!VtoCai), "  /  /    ", rstIvaComp!VtoCai)
            Despacho.Text = IIf(IsNull(rstIvaComp!Despacho), "", rstIvaComp!Despacho)
            
            rstIvaComp.Close
            Call Format_datos
            Call Imprime_Descripcion
    End If
    
    Renglon = 0
        
    For A = 1 To 20
        
            WTipoMovi = "2"
            
            Auxi1 = NroInterno.Text
            Call Ceros(Auxi1, 6)
            XNroInterno = Auxi1
            
            Renglon = A
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            WRenglon = Auxi1$
            
            ClaveImputac = WTipoMovi + XNroInterno + WRenglon
            
            spImputac = "Consultaimputac " + "'" + ClaveImputac + "'"
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
            If rstImputac.RecordCount > 0 Then
            
                If Val(rstImputac!Renglon) <= 10 Then
                    DbGrid1.FirstRow = 0
                    iRow = Val(rstImputac!Renglon) - 1
                        Else
                    DbGrid1.FirstRow = 10
                    iRow = Val(rstImputac!Renglon) - 11
                End If
                        
                DbGrid1.Col = 0
                DbGrid1.Row = iRow
                DbGrid1.Text = rstImputac!Cuenta
                        
                DbGrid1.Col = 2
                DbGrid1.Row = iRow
                DbGrid1.Text = rstImputac!Debito
                DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)

                DbGrid1.Col = 3
                DbGrid1.Row = iRow
                DbGrid1.Text = rstImputac!Credito
                DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                
                WCuenta = rstImputac!Cuenta
                rstImputac.Close
                
                With rstCuenta
                    spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                            DbGrid1.Col = 1
                            DbGrid1.Row = iRow
                            DbGrid1.Text = rstCuenta!Descripcion
                            rstCuenta.Close
                    End If
                End With
                
            End If
            
    Next A
    
    DbGrid1.FirstRow = 0
    DbGrid1.Col = 0
    DbGrid1.Row = 0
    
End Sub


Private Sub cmdAdd_Click()

    ZMes = Mid$(Periodo.Text, 4, 2)
    ZAno = Mid$(Periodo.Text, 7, 4)
    ZEstado = 1

    ZSql = "Select *"
    ZSql = ZSql + " FROM Cierre"
    ZSql = ZSql + " Where Cierre.Mes = " + "'" + ZMes + "'"
    ZSql = ZSql + " and Cierre.Ano = " + "'" + ZAno + "'"
    spCierre = ZSql
    Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
    If rstCierre.RecordCount > 0 Then
        ZEstado = rstCierre!Estado
        rstCierre.Close
    End If
    
    If ZEstado = 1 Then
        m$ = "El mes ya a sido cerrrado, no se puede ingresar ni modificar mas datos"
        A% = MsgBox(m$, 64, "Ingreso de Comprobantes")
        Exit Sub
    End If
        
    If Val(NroInterno.Text) = 0 Then
    
        If Val(WEmpresa) = 1 Then
        
            spIvaComp = "ListaIvacompNumero"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                With rstIvaComp
                    .MoveLast
                    NroInterno.Text = rstIvaComp!NroInterno + 1
                End With
                rstIvaComp.Close
            End If
            m$ = "El numero interno asignado es " + NroInterno.Text
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
                Else
                
            ZHasta = "119000"
            ZSql = ""
            ZSql = ZSql + "Select IvaComp.NroInterno"
            ZSql = ZSql + " FROM Ivacomp"
            ZSql = ZSql + " Where Ivacomp.NroInterno <= " + "'" + ZHasta + "'"
            ZSql = ZSql + " Order by Ivacomp.NroInterno"
            
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                With rstIvaComp
                    .MoveLast
                    NroInterno.Text = rstIvaComp!NroInterno + 1
                End With
                rstIvaComp.Close
            End If
            m$ = "El numero interno asignado es " + NroInterno.Text
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
        End If
        
    End If

    If Val(NroInterno.Text) <> 0 Then
    
        Tipo.Text = TipoComp.ListIndex + 1
            
        WPasa = "S"
        Call Verifica_datos
        
'        With rstProveedor
'            .Index = "Proveedor"
'            .Seek "=", Proveedor.Text
'            If .NoMatch = True Then
'                WPasa = "N"
'                m$ = "Codigo de Proveedor Incorrecto"
'                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
'            End If
'        End With
        
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de emision, formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de vencimiento (1), formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If

        Call Valida_fecha(Vencimiento1.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de vencimiento (2), formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If

        Call Valida_fecha(Periodo.Text, Auxi)
        If Auxi <> "S" Then
            WPasa = "N"
            m$ = "Formato de Fecha de Iva, formato valido : dd/mm/aaaa"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
         
        If Val(Tipo.Text) < 1 Or Val(Tipo.Text) > 3 Then
           WPasa = "N"
           m$ = "Tipo de Comprobante Invalido"
           A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
            
        If Left$(Letra.Text, 1) <> "A" And Left$(Letra.Text, 1) <> "B" And Left$(Letra.Text, 1) <> "C" And Left$(Letra.Text, 1) <> "X" And Left$(Letra.Text, 1) <> "M" And Left$(Letra.Text, 1) <> "I" Then
            WPasa = "N"
            m$ = "Letra del Comprobante Invalido"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        If Pago.ListIndex = 0 Then
            WPasa = "N"
            m$ = "Clausula de Forma de pago no informada"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
        
        If Pago.ListIndex = 2 Then
            If Val(Paridad.Text) = 0 Then
                WPasa = "N"
                m$ = "Paridad no informada"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            End If
        End If
        
        Call Alinea_datos
        ClaveCtaprv = Proveedor.Text + Letra.Text + WTipo + WPunto + WNumero
        spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
        Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
        If RstCtaPrv.RecordCount > 0 Then
            If RstCtaPrv!Saldo <> RstCtaPrv!Total Then
                m$ = "El Comprobante se encuentra total o parcialmente cancelado"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
                WPasa = "N"
            End If
        End If
        
        SumaDebito = 0
        SumaCredito = 0
        
        For A = 0 To 1
        
            Suma = A * 10
            DbGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                        
                DbGrid1.Col = 2
                DbGrid1.Row = iRow
                Debito = DbGrid1.Text
                
                SumaDebito = SumaDebito + Val(Debito)
                        
                DbGrid1.Col = 3
                DbGrid1.Row = iRow
                Credito = DbGrid1.Text
                
                SumaCredito = SumaCredito + Val(Credito)
                
            Next iRow
            
        Next A
                    
        Call Redondeo(SumaDebito)
        Call Redondeo(SumaCredito)
        
        If SumaDebito <> SumaCredito Then
        
            WPasa = "N"
            m$ = "Importe total de debito distinto al del credito"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            
                Else
                
            Dato = Val(Total.Caption)
            Uno = Val(Total.Caption)
            Dato = SumaDebito
            Dato = Pusing("#,###,###.##", Dato)
            Dos = SumaDebito
            
            Call Redondeo(Uno)
            Call Redondeo(Dos)

            If Uno <> Dos Then
                WPasa = "N"
                m$ = "Importe del comprobante distinto a la imputacion contable"
                A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            End If
            
        End If
        
        If WPasa = "S" Then
            Call Alinea_datos
            spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount = 0 Then
                rstIvaComp.Close
                XParam = "'" + Proveedor.Text + "','" _
                             + WTipo + "','" _
                             + WPunto + "','" _
                             + WNumero + "'"
                spIvaComp = "ConsultaIvacompcompro " + XParam
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    rstIvaComp.Close
                    WPasa = "N"
                    m$ = "El comprobante ya se encuentra ingresado"
                    A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
                End If
            End If
        End If
        
        DbGrid1.FirstRow = 0
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        
        If WPasa = "S" Then
    
            Call Alinea_datos

            spIvaComp = "Consultaivacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount = 0 Then
            
                Call Verifica_datos
                
                Rem ALTA DE IVA COMPRAS
                
                XNroInterno = NroInterno.Text
                XProveedor = Proveedor.Text
                XTipo = Tipo.Text
                XLetra = Letra.Text
                XPunto = Punto.Text
                XNumero = Numero.Text
                XFecha = Fecha.Text
                Xvencimiento = Vencimiento.Text
                XVencimiento1 = Vencimiento1.Text
                XPeriodo = Periodo.Text
                XNeto = Neto.Text
                XIva21 = Iva21.Text
                XIva5 = Iva5.Text
                XIva27 = Iva27.Text
                XIb = Ib.Text
                XExento = Exento.Text
                Select Case Val(Tipo.Text)
                    Case 1
                        XImpre = "FC"
                    Case 2
                        XImpre = "ND"
                    Case 3
                        XImpre = "NC"
                        XNeto = Str$(Val(Neto.Text) * -1)
                        XIva21 = Str$(Val(Iva21.Text) * -1)
                        XIva5 = Str$(Val(Iva5.Text) * -1)
                        XIva27 = Str$(Val(Iva27.Text) * -1)
                        XIb = Str$(Val(Ib.Text) * -1)
                        XExento = Str$(Val(Exento.Text) * -1)
                    Case Else
                        XImpre = "  "
                End Select
                XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Contado1.Value = True Then
                    XContado = "1"
                End If
                If Contado2.Value = True Then
                    XContado = "2"
                End If
                XEmpresa = "1"
                XNetolist = ""
                XExentolist = ""
                XParidad = Paridad.Text
                XPAgo = Str$(Pago.ListIndex)
                
                XParam = "'" + XNroInterno + "','" _
                        + XProveedor + "','" + XTipo + "','" _
                        + XLetra + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + Xvencimiento + "','" _
                        + XVencimiento1 + "','" + XPeriodo + "','" _
                        + XNeto + "','" _
                        + XIva21 + "','" _
                        + XIva5 + "','" + XIva27 + "','" _
                        + XIb + "','" + XExento + "','" _
                        + XContado + "','" _
                        + XImpre + "','" + XOrdFecha + "','" _
                        + XEmpresa + "','" + XNetolist + "','" _
                        + XExentolist + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                
                spIvaComp = "AltaIvaCompras " + XParam
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                WNroInterno = NroInterno.Text
                WProveedor = Proveedor.Text
                WTipo = Tipo.Text
                WLetra = Letra.Text
                WPunto = Punto.Text
                WNumero = Numero.Text
                WContado = XContado
                WFecha = Fecha.Text
                Wvencimiento = Vencimiento.Text
                WVencimiento1 = Vencimiento1.Text
                
                    Else
                    
                Call Verifica_datos
                
                Rem modifica DE IVA COMPRAS
                
                XNroInterno = NroInterno.Text
                XProveedor = Proveedor.Text
                XTipo = Tipo.Text
                XLetra = Letra.Text
                XPunto = Punto.Text
                XNumero = Numero.Text
                XFecha = Fecha.Text
                Xvencimiento = Vencimiento.Text
                XVencimiento1 = Vencimiento1.Text
                XPeriodo = Periodo.Text
                XNeto = Neto.Text
                XIva21 = Iva21.Text
                XIva5 = Iva5.Text
                XIva27 = Iva27.Text
                XIb = Ib.Text
                XExento = Exento.Text
                Select Case Val(Tipo.Text)
                    Case 1
                        XImpre = "FC"
                    Case 2
                        XImpre = "ND"
                    Case 3
                        XImpre = "NC"
                        XNeto = Str$(Val(Neto.Text) * -1)
                        XIva21 = Str$(Val(Iva21.Text) * -1)
                        XIva5 = Str$(Val(Iva5.Text) * -1)
                        XIva27 = Str$(Val(Iva27.Text) * -1)
                        XIb = Str$(Val(Ib.Text) * -1)
                        XExento = Str$(Val(Exento.Text) * -1)
                    Case Else
                        XImpre = "  "
                End Select
                XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                If Contado1.Value = True Then
                    XContado = "1"
                End If
                If Contado2.Value = True Then
                    XContado = "2"
                End If
                XEmpresa = "1"
                XNetolist = ""
                XExentolist = ""
                XParidad = Paridad.Text
                XPAgo = Str$(Pago.ListIndex)
                
                XParam = "'" + XNroInterno + "','" _
                        + XProveedor + "','" + XTipo + "','" _
                        + XLetra + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + Xvencimiento + "','" _
                        + XVencimiento1 + "','" + XPeriodo + "','" _
                        + XNeto + "','" _
                        + XIva21 + "','" _
                        + XIva5 + "','" + XIva27 + "','" _
                        + XIb + "','" + XExento + "','" _
                        + XContado + "','" _
                        + XImpre + "','" + XOrdFecha + "','" _
                        + XEmpresa + "','" + XNetolist + "','" _
                        + XExentolist + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                
                spIvaComp = "ActualizaIvaCompras " + XParam
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                WNroInterno = NroInterno.Text
                WProveedor = Proveedor.Text
                WTipo = Tipo.Text
                WLetra = Letra.Text
                WPunto = Punto.Text
                WNumero = Numero.Text
                WContado = XContado
                WFecha = Fecha.Text
                Wvencimiento = Vencimiento.Text
                WVencimiento1 = Vencimiento1.Text
                
            End If
            
            XParam = "'" + XNroInterno + "','" _
                         + Cai.Text + "','" _
                         + VtoCai.Text + "'"
                
            spIvaComp = "ActualizaIvaComprasCai " + XParam
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE IvaComp SET "
            ZSql = ZSql + " Despacho = " + "'" + Despacho.Text + "'"
            ZSql = ZSql + " Where NroInterno = " + "'" + XNroInterno + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
            Rem borra LAS IMPUTACIONES CONTABLES
        
            spImputac = "BorrarImputac " + "'" + NroInterno.Text + "'"
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
        
            Rem GRABA LAS IMPUTACIONES CONTABLES
        
            Renglon = 0
            Auxi1 = WNroInterno
            Call Ceros(Auxi1, 6)
            WNroInterno = Auxi1
            
                                        
            For A = 0 To 1
        
                Suma = A * 10
                DbGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRow = iRow
                        
                    DbGrid1.Col = 0
                    DbGrid1.Row = iRow
                    WCuenta = DbGrid1.Text
                        
                    DbGrid1.Col = 2
                    DbGrid1.Row = iRow
                    Debito = Val(DbGrid1.Text)
                        
                    DbGrid1.Col = 3
                    DbGrid1.Row = iRow
                    Dato = DbGrid1.Text
                    Credito = Val(DbGrid1.Text)
                
                    If WCuenta <> "" Then
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        
                        XTipomovi = "2"
                        XNroInterno = WNroInterno
                        XProveedor = WProveedor
                        XTipocomp = WTipo
                        XLetracomp = WLetra
                        XPuntocomp = WPunto
                        XNrocomp = WNumero
                        XRenglon = Str$(Renglon + 1)
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        XRenglon = Auxi1$
                        XFecha = Fecha.Text
                        XObservaciones = ""
                        XCuenta = WCuenta
                        If Debito <> "" Then
                            XDebito = Str$(Debito)
                        End If
                        If Credito <> "" Then
                            XCredito = Str$(Credito)
                        End If
                        XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        XTitulo = "Compras"
                        XEmpresa = "1"
                        XClave = XTipomovi + XNroInterno + XRenglon
                        XDebitolist = ""
                        XCreditolist = ""
                        
                        XParam = "'" + XClave + "','" _
                                + XTipomovi + "','" + XProveedor + "','" _
                                + XTipocomp + "','" _
                                + XLetracomp + "','" + XPuntocomp + "','" _
                                + XNrocomp + "','" _
                                + XRenglon + "','" _
                                + XFecha + "','" + XObservaciones + "','" _
                                + XCuenta + "','" _
                                + XDebito + "','" _
                                + XCredito + "','" + XFechaOrd + "','" _
                                + XTitulo + "','" + XEmpresa + "','" _
                                + XDebitolist + "','" _
                                + XCreditolist + "','" _
                                + XNroInterno + "'"
                                
                        spImputac = "AltaImputacion " + XParam
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                                        
                Next iRow
            
            Next A
        
            Rem graba la cta.cte

            If Val(WContado) = 2 Then
        
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount = 0 Then
                
                    XProveedor = WProveedor
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = WFecha
                    XEstado = "1"
                    Xvencimiento = Wvencimiento
                    XVencimiento1 = WVencimiento1
                    XNroInterno = WNroInterno
                    XTotal = Total.Caption
                    XSaldo = Total.Caption
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
                    XOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                    Select Case Val(WTipo)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XTotal = Str$(Val(Total.Caption) * -1)
                            XSaldo = Str$(Val(Total.Caption) * -1)
                        Case Else
                            XImpre = ""
                    End Select
                    XEmpresa = "1"
                    XSaldolist = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = Paridad.Text
                    XPAgo = Str$(Pago.ListIndex)
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spConsulta = "AltaCtaPrv " + XParam
                    Set rstConsulta = db.OpenRecordset(spConsulta + cParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    RstCtaPrv.Close
                  
                    XProveedor = WProveedor
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = WFecha
                    XEstado = "1"
                    Xvencimiento = Wvencimiento
                    XVencimiento1 = WVencimiento1
                    XNroInterno = WNroInterno
                    XTotal = Total.Caption
                    XSaldo = Total.Caption
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
                    XOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                    Select Case Val(WTipo)
                        Case 1
                            XImpre = "FC"
                        Case 2
                            XImpre = "ND"
                        Case 3
                            XImpre = "NC"
                            XTotal = Str$(Val(Total.Caption) * -1)
                            XSaldo = Str$(Val(Total.Caption) * -1)
                        Case Else
                            XImpre = ""
                    End Select
                    XEmpresa = "1"
                    XSaldolist = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = Paridad.Text
                    XPAgo = Str$(Pago.ListIndex)
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                    
                    spCtaprv = "ModificaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            End If
            
            Sql1 = "UPDATE Proveedor SET "
            Sql2 = " Cai = " + "'" + Cai.Text + "',"
            Sql3 = " VtoCai = " + "'" + VtoCai.Text + "'"
            Sql4 = " Where Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = Sql1 + Sql2 + Sql3 + Sql4
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            Call CmdLimpiar_Click
        
        End If
        
        NroInterno.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    WPasa = "S"

    Call Alinea_datos
    ClaveCtaprv = Proveedor.Text + Letra.Text + WTipo + WPunto + WNumero
    spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
    If RstCtaPrv.RecordCount > 0 Then
        If RstCtaPrv!Saldo <> RstCtaPrv!Total Then
            m$ = "El Comprobante se encuentra total o parcialmente cancelado"
            A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            WPasa = "N"
        End If
    End If
    
    If WPasa = "S" Then

        spIvaComp = "ConsultaIvacomp " + "'" + NroInterno.Text + "'"
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                spIvaComp = "BorrarIvacomp " + "'" + NroInterno.Text + "'"
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                If Val(NroInterno.Text) <> 0 Then
                    ZSql = ""
                    ZSql = ZSql + "DELETE CtaCtePrv"
                    ZSql = ZSql + " Where NroInterno = " + "'" + NroInterno.Text + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
                spImputac = "BorrarImputac " + "'" + NroInterno.Text + "'"
                Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                    
                Call CmdLimpiar_Click
                    
            End If
        End If
        
    End If
    
    NroInterno.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    For A = 0 To 1
    Suma = A * 10
    DbGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 3
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    DbGrid1.FirstRow = 0

    NroInterno.Text = ""
    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Vencimiento1.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Ib.Text = ""
    Exento.Text = ""
    Total.Caption = ""
    Paridad.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    Despacho.Text = ""
    Contado1.Value = False
    Contado2.Value = True
    DesProveedor.Caption = ""
    TipoComp.ListIndex = 0
    Pago.ListIndex = 0
    
    NroInterno.Text = ""
    spIvaComp = "ListaIvacompNumero"
    Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstIvacomp.RecordCount > 0 Then
    Rem     With rstIvacomp
    Rem         .MoveLast
    Rem         NroInterno.Text = rstIvacomp!NroInterno + 1
    Rem     End With
    Rem     rstIvacomp.Close
    Rem End If
    
    NroInterno.SetFocus
End Sub

Private Sub cmdClose_Click()

    CmdLimpiar_Click

    NroInterno.SetFocus
    PrgCompras.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub



Private Sub Proveedor_KeyPress(KeyAscii As Integer)

    WProveedor = Proveedor.Text
    Proveedor.Text = WProveedor

    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = RstProveedor!Nombre
        If Trim(Cai.Text) = "" Then
            Cai.Text = IIf(IsNull(RstProveedor!Cai), "", RstProveedor!Cai)
            VtoCai.Text = IIf(IsNull(RstProveedor!VtoCai), "  /  /    ", RstProveedor!VtoCai)
        End If
        Letra.SetFocus
        RstProveedor.Close
            Else
        Proveedor.SetFocus
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Letra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Left$(Letra.Text, 1) = "A" Or Left$(Letra.Text, 1) = "B" Or Left$(Letra.Text, 1) = "C" Or Left$(Letra.Text, 1) = "X" Or Left$(Letra.Text, 1) = "M" Or Left$(Letra.Text, 1) = "I" Then
            Punto.SetFocus
                Else
            Letra.SetFocus
        End If
    End If
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
        Punto.Text = WPunto
        Numero.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WNumero = Numero.Text
        Call Ceros(WNumero, 8)
        Numero.Text = WNumero
        Cai.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoCai.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub VtoCai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(VtoCai.Text, Auxi)
        If Auxi = "S" Or VtoCai.Text = "  /  /    " Then
            WNumero = Numero.Text
            Call Ceros(WNumero, 8)
            Numero.Text = WNumero
            Fecha.SetFocus
                Else
            VtoCai.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Periodo.Text = "  /  /    " Then
                Periodo.Text = Fecha.Text
            End If
            Vencimiento.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub



Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Wvencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
            If Wvencimiento >= WFecha Then
                Vencimiento1.SetFocus
                    Else
                Vencimiento.SetFocus
            End If
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Vencimiento1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento1.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Wvencimiento = Right$(Vencimiento1.Text, 4) + Mid$(Vencimiento1.Text, 4, 2) + Left$(Vencimiento1.Text, 2)
            If Wvencimiento >= WFecha Then
                Periodo.SetFocus
                    Else
                Vencimiento1.SetFocus
            End If
                Else
            Vencimiento1.SetFocus
        End If
    End If
End Sub

Private Sub Periodo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Periodo.Text, Auxi)
        If Auxi = "S" Then
            Pago.SetFocus
                Else
            Periodo.SetFocus
        End If
    End If
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Pago.ListIndex = 2 Then
            Paridad.SetFocus
                Else
            Paridad.Text = ""
            Neto.SetFocus
        End If
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = Pusing("#,###.####", Paridad.Text)
        Neto.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Neto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Neto.Text = Pusing("#,###,###.##", Neto.Text)
        Call Calcula_total
        Iva21.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
        Call Calcula_total
        Iva5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva5.Text = Pusing("#,###,###.##", Iva5.Text)
        Call Calcula_total
        Iva27.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva27_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva27.Text = Pusing("#,###,###.##", Iva27.Text)
        Call Calcula_total
        Ib.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ib_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ib.Text = Pusing("#,###,###.##", Ib.Text)
        Call Calcula_total
        Exento.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Exento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Exento.Text = Pusing("#,###,###.##", Exento.Text)
        Call Calcula_total
        Despacho.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Despacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroInterno_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(NroInterno.Text) <> 0 Then
        
            XNroInterno = NroInterno.Text
            XProveedor = Proveedor.Text
            Rem XTipo = Tipo.Text
            XLetra = Letra.Text
            XPunto = Punto.Text
            XNumero = Numero.Text
            WNumero = Numero.Text
            Call Ceros(WNumero, 8)
            Numero.Text = WNumero
            
            spIvaComp = "ConsultaIvacomp " + "'" + NroInterno.Text + "'"
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
                    NroInterno.Text = XNroInterno
                    Proveedor.Text = XProveedor
                    Rem Tipo.Text = XTipo
                    Letra.Text = XLetra
                    Punto.Text = XPunto
                    Numero.Text = XNumero
                    '
                    rstIvaComp.Close
                    Call Imprime_Datos
                    '
                    Existe = "S"
                        Else
                    CmdLimpiar_Click
                    NroInterno.Text = XNroInterno
                    Proveedor.Text = XProveedor
                    Rem Tipo.Text = XTipo
                    Letra.Text = XLetra
                    Punto.Text = XPunto
                    Numero.Text = XNumero
                    Existe = "N"
                    Call Imprime_Descripcion
            End If
            
        End If
        
        Proveedor.SetFocus
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Cuentas Contables"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spProveedor = "ListaProveedoresOrdConsulta"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
        
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "      " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
            
            End If
            
        Case 1
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
            
            With rstCuenta
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCuenta!Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCuenta.Close
            
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                    Proveedor.Text = RstProveedor!Proveedor
                    Cai.Text = IIf(IsNull(RstProveedor!Cai), "", RstProveedor!Cai)
                    Cai.Text = Trim(Cai.Text)
                    VtoCai.Text = IIf(IsNull(RstProveedor!VtoCai), "  /  /    ", RstProveedor!VtoCai)
                    RstProveedor.Close
                    Call Imprime_Descripcion
                        Else
                    CmdLimpiar_Click
                    Proveedor.Text = Claveven$
            End If
            Proveedor.SetFocus
            
        Case 1
            Indice = Pantalla.ListIndex
            WCuenta = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                    DbGrid1.Col = 0
                    DbGrid1.Text = rstCuenta!Cuenta
                    DbGrid1.Col = 1
                    DbGrid1.Text = rstCuenta!Descripcion
                    DbGrid1.Col = 2
                    KeyCode = 0
                    DbGrid1.SetFocus
                    rstCuenta.Close
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
            Case 0
                If KeyCode = 13 Then
                    spCuenta = "ConsultaCuentas " + "'" + DbGrid1.Text + "'"
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                            DbGrid1.Col = 1
                            DbGrid1.Text = rstCuenta!Descripcion
                            DbGrid1.Col = 2
                            KeyCode = 0
                            rstCuenta.Close
                                Else
                            DbGrid1.Col = 0
                            KeyCode = 0
                    End If
                End If
                
            Case 2
                If KeyCode = 13 Then
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    DbGrid1.Col = 3
                    KeyCode = 0
                End If
                
            Case 3
                If KeyCode = 13 Then
                    DbGrid1.Text = Pusing("#,###,###.##", DbGrid1.Text)
                    If DbGrid1.Row < 21 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
            
            Case Else
            
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
        UserData(iCol, mTotalRows - 1) = DbGrid1.Columns(iCol).DefaultValue
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
ReDim UserData(0 To 3, 0 To 20)

mTotalRows& = 20

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DbGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DbGrid1.Columns.Count - 1 To 0 Step -1
     DbGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 3
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Cuenta Contable"
             DbGrid1.Columns(newcnt).Width = 1600
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Nombre"
             DbGrid1.Columns(newcnt).Width = 3500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Debito"
             DbGrid1.Columns(newcnt).Width = 1500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Credito"
             DbGrid1.Columns(newcnt).Width = 1500
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = False
             DbGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    
    NroInterno.Text = ""
    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Vencimiento1.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Ib.Text = ""
    Exento.Text = ""
    Total.Caption = ""
    Contado1.Value = False
    Contado2.Value = True
    DesProveedor.Caption = ""
    Paridad.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    Despacho.Text = ""
    
    NroInterno.Text = ""
    Rem spIvacomp = "ListaIvacompNumero"
    Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstIvacomp.RecordCount > 0 Then
    Rem     With rstIvacomp
    Rem         .MoveLast
    Rem         NroInterno.Text = rstIvacomp!NroInterno + 1
    Rem     End With
    Rem     rstIvacomp.Close
    Rem End If
    
    TipoComp.Clear
    
    TipoComp.AddItem "Factura"
    TipoComp.AddItem "N.Debito"
    TipoComp.AddItem "N.Credito"
    
    TipoComp.ListIndex = 0
    
    Pago.Clear
    
    Pago.AddItem ""
    Pago.AddItem "Pesos"
    Pago.AddItem "Clausula Dolar"
    
    Pago.ListIndex = 0
    

End Sub



