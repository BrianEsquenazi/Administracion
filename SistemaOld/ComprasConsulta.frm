VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgComprasConsulta 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Comprobantes de Proveedores"
   ClientHeight    =   8235
   ClientLeft      =   375
   ClientTop       =   435
   ClientWidth     =   11190
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   11190
   Begin VB.TextBox Remito 
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
      MaxLength       =   30
      TabIndex        =   63
      Text            =   " "
      Top             =   960
      Width           =   2535
   End
   Begin VB.CheckBox Rechazado 
      Caption         =   "Ch.Rech."
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
      Left            =   7320
      TabIndex        =   61
      Top             =   480
      Width           =   1215
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5520
      Width           =   375
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
      TabIndex        =   56
      Top             =   4440
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   55
      Top             =   5040
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
      TabIndex        =   54
      Top             =   4440
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   5160
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
      TabIndex        =   52
      Top             =   5040
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   5160
      Width           =   375
   End
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
      TabIndex        =   48
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
      Left            =   9480
      MaxLength       =   14
      TabIndex        =   44
      Text            =   " "
      Top             =   120
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
      TabIndex        =   43
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
      TabIndex        =   41
      Top             =   1320
      Width           =   1815
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      Left            =   3720
      TabIndex        =   23
      Top             =   2880
      Width           =   3495
      Begin VB.OptionButton Contado3 
         Caption         =   "Nacion"
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
         Left            =   2280
         TabIndex        =   64
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Contado2 
         Caption         =   "Cta.Cte."
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
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1095
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
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
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
      Height          =   825
      Left            =   8400
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10200
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox VtoCai 
      Height          =   285
      Left            =   9840
      TabIndex        =   46
      Top             =   480
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3975
      Left            =   120
      TabIndex        =   50
      Top             =   3720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7011
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3000
      TabIndex        =   57
      Top             =   4440
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
   Begin VB.Label Label21 
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
      Left            =   6720
      TabIndex        =   62
      Top             =   960
      Width           =   975
   End
   Begin VB.Label TotalDebito 
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
      Left            =   6600
      TabIndex        =   60
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label TotalCredito 
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
      Left            =   8280
      TabIndex        =   59
      Top             =   7680
      Width           =   1575
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
      TabIndex        =   49
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
      Left            =   8760
      TabIndex        =   47
      Top             =   480
      Width           =   855
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
      Left            =   8760
      TabIndex        =   45
      Top             =   120
      Width           =   615
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
      TabIndex        =   42
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
      TabIndex        =   40
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
Attribute VB_Name = "PrgComprasConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim ZNroRemito(100) As String
Dim ZNroOrden(100, 2) As String
Dim ZZRemito As String
Dim EmpresaTrabajo As String
Dim ZPyme As String

Dim ZZCuotas As Integer
Dim ZZMesCuota As Integer
Dim ZZAnoCuota As Integer
Dim ZZValorCuota As Double
Dim ZZIva As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

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
            Contado3.Value = False
            Select Case Val(rstIvaComp!Contado)
                Case 1
                    Contado1.Value = True
                Case 2
                    Contado2.Value = True
                Case 3
                    Contado3.Value = True
                Case Else
            End Select
            Paridad.Text = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
            Pago.ListIndex = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
            Cai.Text = IIf(IsNull(rstIvaComp!Cai), "", rstIvaComp!Cai)
            Cai.Text = Trim(Cai.Text)
            VtoCai.Text = IIf(IsNull(rstIvaComp!VtoCai), "  /  /    ", rstIvaComp!VtoCai)
            Despacho.Text = IIf(IsNull(rstIvaComp!Despacho), "", rstIvaComp!Despacho)
            Remito.Text = IIf(IsNull(rstIvaComp!Remito), "", rstIvaComp!Remito)
            ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
            Rechazado.Value = ZRechazado
            
            rstIvaComp.Close
            Call Format_datos
            Call Imprime_Descripcion
    End If
    
    Renglon = 0
    Call Limpia_Vector
        
    For A = 1 To 50
        
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
            
            WVector1.Row = Val(rstImputac!Renglon)
                
            WVector1.Col = 1
            WVector1.Text = rstImputac!Cuenta
                        
            WVector1.Col = 3
            WVector1.Text = Str$(rstImputac!Debito)
            WVector1.Text = Pusing("#,###,###.##", WVector1.Text)

            WVector1.Col = 4
            WVector1.Text = Str$(rstImputac!Credito)
            WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                
            WCuenta = rstImputac!Cuenta
            rstImputac.Close
                
            With rstCuenta
                spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End With
                
        End If
            
    Next A
    
    Call Calcula_Click
    
    WVector1.Col = 1
    WVector1.Row = 1
    
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
                Periodo.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            End If
            
            XEmpresa = WEmpresa
                        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Claveven$ = Proveedor.Text
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                  
                Tipom = RstProveedor!TipoProv
                Dias = RstProveedor!Dias
                Dias = UCase(Dias)
                If Val(Dias) = 0 Then
                  Vencimiento.Text = Fecha.Text
                  Dias = 0
                End If
                Dias = Left$(Dias, 3)
                
                If Dias = "CON" Then
                    Vencimiento.Text = Fecha.Text
                     Else
                  If Tipom = 1 Then
                  Rem   Dias = "60"
                      Else
                       If Dias <> "15" Or Dias <> "30" Then
                    Rem     Dias = "30"
                         End If
                     End If
                End If
                
                
                If Dias <> "CON" Then
                    ZZDias = Trim(Str$(Val(Dias)))
                    Fecha2 = DateValue(Fecha.Text)
                    Vencimiento.Text = DateAdd("d", ZZDias, Fecha2)
                End If
                    
                RstProveedor.Close
                    
            End If
            
            Call Conecta_Empresa
            
            Rem fin by nan
            
            Vencimiento.SetFocus
            Vencimiento1.Text = Vencimiento.Text
            
           Rem ene end by nan
       Rem     Remito.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Remito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Rem Vencimiento.SetFocus
          Pago.SetFocus
          
        If Trim(Remito.Text) <> "" Then
            Call Remito_dblclick
        End If
    End If
End Sub

Private Sub Remito_dblclick()


    ZZPasaRemito = Remito.Text
    ZZPasaProveedor = Proveedor.Text
    ZZPasaProceso = 0
    
    Call Verifica_Pyme
    
    If ZPyme = "S" Then
        If Val(Cuotas.Text) = 0 Then
            Cuotas.Text = Trim(Str$(ZZCuotas))
        End If
        If Val(MesCuota.Text) = 0 Or Val(AnoCuota.Text) = 0 Then
            MesCuota.Text = Trim(Str$(ZZMesCuota))
            AnoCuota.Text = Trim(Str$(ZZAnoCuota))
        End If
        Contado3.Value = True
    End If
    
    PrgConsultaInforme.Show

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
          Remito.SetFocus
          Rem  Pago.SetFocus
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
        If Val(Iva21.Text) = 0 Then
            If Letra.Text = "A" Or Letra.Text = "M" Then
                ZZIva = Val(Neto.Text) * 0.21
                Call Redondeo(ZZIva)
                Iva21.Text = Str$(ZZIva)
                Iva21.Text = Pusing("#,###,###.##", Iva21.Text)
            End If
        End If
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
        Entra = "S"
        For iRow = 1 To 50
            ZZCuenta = Trim(WVector1.TextMatrix(iRow, 1))
            ZZDebito = WVector1.TextMatrix(iRow, 3)
            ZZCredito = WVector1.TextMatrix(iRow, 4)
            If Trim(ZZCuenta) <> "" Or Val(ZZDebito) <> 0 Or Val(ZZCredito) <> 0 Then
                Entra = "N"
            End If
        Next iRow
        If Entra = "S" Then
        
            ZZLugar = 0
            
            If Val(Total.Caption) <> 0 Then
                ZZLugar = ZZLugar + 1
                If Letra.Text = "I" Then
                    WVector1.TextMatrix(ZZLugar, 1) = "2010"
                        Else
                    WVector1.TextMatrix(ZZLugar, 1) = "2001"
                End If
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Total.Caption)
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Total.Caption)
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva21.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "151"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva21.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva21.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva27.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "151"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva27.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva27.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Iva5.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "152"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Iva5.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Iva5.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
            If Val(Ib.Text) <> 0 Then
                ZZLugar = ZZLugar + 1
                WVector1.TextMatrix(ZZLugar, 1) = "164"
                WVector1.TextMatrix(ZZLugar, 2) = ""
                If TipoComp.ListIndex <> 2 Then
                    WVector1.TextMatrix(ZZLugar, 3) = Pusing("#,###,###.##", Ib.Text)
                    WVector1.TextMatrix(ZZLugar, 4) = ""
                        Else
                    WVector1.TextMatrix(ZZLugar, 4) = Pusing("#,###,###.##", Ib.Text)
                    WVector1.TextMatrix(ZZLugar, 3) = ""
                End If
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(ZZLugar, 1) + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.TextMatrix(ZZLugar, 2) = rstCuenta!Descripcion
                    rstCuenta.Close
                End If
            End If
            
        End If
        
        Call Calcula_Click
        
        WVector1.Col = 1
        WVector1.Row = ZZLugar + 1
        Call StartEdit
        
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
Rem by nan
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
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            WCuenta = WIndice.List(Indice)
            spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstCuenta!Cuenta
                WVector1.Col = 2
                WVector1.Text = rstCuenta!Descripcion
                WVector1.Col = 3
                Call StartEdit
                rstCuenta.Close
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

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
    Contado3.Value = False
    DesProveedor.Caption = ""
    Paridad.Text = ""
    Cai.Text = ""
    VtoCai.Text = "  /  /    "
    Despacho.Text = ""
    Remito.Text = ""
    
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
    
    NroInterno.Text = WPasaNroInterno
    Call NroInterno_Keypress(13)

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
        Call Calcula_Click
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
            End If
            Call StartEdit
    
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
        Case 4
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
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstCuenta!Descripcion
                WVector1.Col = 2
                rstCuenta.Close
                    Else
                WControl = "N"
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 50 To 1 Step -1
        
        ZLegajo = WVector1.TextMatrix(iRow, 1)
            
        If ZLegajo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 0 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
    Call Calcula_Click
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem ""
    Opcion.AddItem ""

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
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
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 51
    
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Cuenta"
                WVector1.ColWidth(Ciclo) = 1700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 4500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Debito"
                WVector1.ColWidth(Ciclo) = 1700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 4
                WVector1.Text = "Credito"
                WVector1.ColWidth(Ciclo) = 1700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
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

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Calcula_Click()

    SumaDebito = 0
    SumaCredito = 0
        
    For iRow = 1 To 50
        Debito = WVector1.TextMatrix(iRow, 3)
        SumaDebito = SumaDebito + Val(Debito)
        Credito = WVector1.TextMatrix(iRow, 4)
        SumaCredito = SumaCredito + Val(Credito)
    Next iRow
                    
    Call Redondeo(SumaDebito)
    Call Redondeo(SumaCredito)
    
    TotalDebito.Caption = Str$(SumaDebito)
    TotalCredito.Caption = Str$(SumaCredito)
    
    TotalDebito.Caption = Pusing("#,###,###.##", TotalDebito.Caption)
    TotalCredito.Caption = Pusing("#,###,###.##", TotalCredito.Caption)

End Sub

Private Sub CerrarBusquedaNro_Click()
    BusquedaNro.Visible = False
End Sub

Private Sub ConsultaII_Click()

    ProveedorII.Text = ""
    DesProveedorII.Caption = ""
    LetraII.Text = ""
    NumeroII.Text = ""
    PuntoII.Text = ""
    TipoII.Text = ""
    TipoCompII.ListIndex = 0
    
    BusquedaNro.Visible = True
    
    ProveedorII.SetFocus

End Sub


Private Sub ProveedorII_KeyPress(KeyAscii As Integer)

    WProveedor = ProveedorII.Text
    ProveedorII.Text = WProveedor

    spProveedor = "ConsultaProveedores " + "'" + ProveedorII.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        DesProveedorII.Caption = RstProveedor!Nombre
        LetraII.SetFocus
        RstProveedor.Close
            Else
        ProveedorII.SetFocus
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub LetraII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Left$(LetraII.Text, 1) = "A" Or Left$(LetraII.Text, 1) = "B" Or Left$(LetraII.Text, 1) = "C" Or Left$(LetraII.Text, 1) = "X" Or Left$(LetraII.Text, 1) = "M" Or Left$(LetraII.Text, 1) = "I" Then
            PuntoII.SetFocus
                Else
            LetraII.SetFocus
        End If
    End If
End Sub

Private Sub PuntoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = PuntoII.Text
        Call Ceros(WPunto, 4)
        PuntoII.Text = WPunto
        NumeroII.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NumeroII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        TipoII.Text = TipoCompII.ListIndex + 1
        WTipo = TipoII.Text
        Call Ceros(WTipo, 2)
        
        WPunto = PuntoII.Text
        Call Ceros(WPunto, 4)
        
        WNumero = NumeroII.Text
        Call Ceros(WNumero, 8)
        
        ZSql = "Select *"
        ZSql = ZSql + " FROM Ivacomp"
        ZSql = ZSql + " Where Ivacomp.Proveedor = " + "'" + ProveedorII.Text + "'"
        ZSql = ZSql + " and Ivacomp.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Ivacomp.Letra = " + "'" + LetraII.Text + "'"
        ZSql = ZSql + " and Ivacomp.Punto = " + "'" + WPunto + "'"
        ZSql = ZSql + " and Ivacomp.Numero = " + "'" + WNumero + "'"
        spIvaComp = ZSql
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            NroInterno.Text = rstIvaComp!NroInterno
            rstIvaComp.Close
            BusquedaNro.Visible = False
            NroInterno_Keypress (13)
                Else
            m$ = "Factura no ingresada"
            A% = MsgBox(m$, 64, "Ingreso de Comprobantes")
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Verifica_Pyme()
    
    ZZCuotas = 0
    ZZMesCuota = 0
    ZZAnoCuota = 0
    
    Erase ZNroRemito
    Erase ZNroOrden
    ZLugarII = 0
    ZLugarIII = 0
    
    ZPyme = "N"
    
    ZZCargaRemito = Trim(ZZPasaRemito)
    
    If Val(ZZCargaRemito) = 0 Then
        Exit Sub
    End If
    
    Do
        MyPos = InStr(ZZCargaRemito, ",")
        If MyPos = 0 Then
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = ZZCargaRemito
            Exit Do
                Else
            ZLugarII = ZLugarII + 1
            ZNroRemito(ZLugarII) = Mid$(ZZCargaRemito, 1, MyPos - 1)
            ZZCargaRemito = Mid$(ZZCargaRemito, MyPos + 1, 100)
        End If
    Loop
    
    Call Busca_Empresa
    XEmpresa = WEmpresa
    
    Select Case Val(EmpresaTrabajo)
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
    
    For CicloII = 1 To ZLugarII
    
        ZZRemito = ZNroRemito(CicloII)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Informe.Remito = " + "'" + ZZRemito + "'"
        ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            With rstInforme
                .MoveFirst
                Do
                    If .EOF = False Then
                        
                        If rstInforme!Cantidad <> 0 Then
                            WLugarIII = WLugarIII + 1
                            ZNroOrden(WLugarIII, 1) = Str$(rstInforme!Orden)
                            ZNroOrden(WLugarIII, 2) = rstInforme!Articulo
                        End If
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstInforme.Close
        End If
        
    Next CicloII
    
    For Ciclo = 1 To WLugarIII
    
        ZZOrden = ZNroOrden(Ciclo, 1)
        ZZArticulo = ZNroOrden(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZZOrden + "'"
        ZSql = ZSql + " and Orden.Articulo = " + "'" + ZZArticulo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            ZTarjeta = IIf(IsNull(rstOrden!Tarjeta), "0", rstOrden!Tarjeta)
            If ZTarjeta = 1 Then
                ZPyme = "S"
                
                ZZCuotas = IIf(IsNull(rstOrden!Cuotas), "", rstOrden!Cuotas)
                ZZMesCuota = IIf(IsNull(rstOrden!MesCuota), "", rstOrden!MesCuota)
                ZZAnoCuota = IIf(IsNull(rstOrden!AnoCuota), "", rstOrden!AnoCuota)
                
            End If
            rstOrden.Close
        End If
    
    Next Ciclo
        
    Call Conecta_Empresa
    
End Sub



Private Sub Busca_Empresa()

    EmpresaTrabajo = 0
    EmpresaAnterior = WEmpresa
    XEmpresa = WEmpresa
    
    If EmpresaAnterior = 1 Then

        For Va = 1 To 7
    
            Select Case Va
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
                Case 7
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = WEmpresa
                rstInforme.Close
                Exit For
            End If
        
        Next Va
        
            Else
        
        For Va = 1 To 4
    
            Select Case Va
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
                Case 4
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Remito = " + "'" + ZNroRemito(1) + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                EmpresaTrabajo = WEmpresa
                rstInforme.Close
            End If
        
        Next Va
        
    End If
    
    Call Conecta_Empresa
    
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

Private Sub Contado3_Click()
    PantaPyme.Visible = True
    Cuotas.SetFocus
End Sub

Private Sub CierraPyme_Click()
    PantaPyme.Visible = False
    Pago.SetFocus
End Sub

