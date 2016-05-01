VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactupOLD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturacion de Pedidos en $"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      Height          =   2655
      Left            =   5400
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox WCanti5 
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
         TabIndex        =   68
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox WCanti4 
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
         TabIndex        =   67
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WLote5 
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
         Left            =   240
         TabIndex        =   66
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox WLote4 
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
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox WCanti3 
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
         TabIndex        =   64
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WCanti2 
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
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCanti1 
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
         TabIndex        =   62
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Wlote3 
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
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WLote2 
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
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WLote1 
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
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
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
         Left            =   1440
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partida"
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
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ComboBox Tipoventa 
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
      Left            =   3240
      TabIndex        =   55
      Top             =   1200
      Width           =   2655
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
      Height          =   495
      Left            =   10200
      TabIndex        =   54
      Top             =   720
      Width           =   1215
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
      TabIndex        =   48
      Text            =   " "
      Top             =   7440
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
      TabIndex        =   47
      Text            =   " "
      Top             =   7080
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
      TabIndex        =   46
      Text            =   " "
      Top             =   6720
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
      TabIndex        =   45
      Text            =   " "
      Top             =   6360
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
      TabIndex        =   44
      Text            =   " "
      Top             =   6000
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
      TabIndex        =   43
      Text            =   " "
      Top             =   7440
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
      TabIndex        =   42
      Text            =   " "
      Top             =   7080
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
      TabIndex        =   41
      Text            =   " "
      Top             =   6720
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
      TabIndex        =   40
      Text            =   " "
      Top             =   6360
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
      TabIndex        =   39
      Text            =   " "
      Top             =   6000
      Width           =   975
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   34
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
      Height          =   495
      Left            =   9120
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   8760
      TabIndex        =   23
      Top             =   5760
      Width           =   2535
      Begin VB.Label Label18 
         Caption         =   "Ing.Brutos"
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
         TabIndex        =   70
         Top             =   960
         Width           =   975
      End
      Begin VB.Label ImpoIb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   69
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Dto."
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
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Dto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Interes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   31
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Iva2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Iva1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   1215
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1680
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
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
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
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
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
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
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
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   22
      Text            =   " "
      Top             =   120
      Width           =   1335
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
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   720
      Width           =   975
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
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Orden 
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
      MaxLength       =   10
      TabIndex        =   18
      Text            =   " "
      Top             =   1200
      Width           =   1335
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
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   16
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1800
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
      TabIndex        =   11
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3960
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   7
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
      Height          =   450
      Left            =   8040
      TabIndex        =   5
      Top             =   120
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
      Height          =   450
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Width           =   975
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
      Height          =   450
      Left            =   10200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   1320
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
      Height          =   1980
      ItemData        =   "PrgfactupOLD.frx":0000
      Left            =   6480
      List            =   "PrgfactupOLD.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "PrgfactupOLD.frx":0015
      TabIndex        =   2
      Top             =   1680
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      TabIndex        =   53
      Top             =   7440
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
      TabIndex        =   52
      Top             =   7080
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
      TabIndex        =   51
      Top             =   6720
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
      TabIndex        =   50
      Top             =   6360
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
      TabIndex        =   49
      Top             =   6000
      Width           =   855
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
      Left            =   5640
      TabIndex        =   33
      Top             =   840
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Orden de compra"
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
      Top             =   1200
      Width           =   1815
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
      Left            =   3240
      TabIndex        =   15
      Top             =   840
      Width           =   1095
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
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   4695
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
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Factura"
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
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgFactupOLD"
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
Private WPlazo1 As Integer
Private WPlazo2 As Integer
Private WDias1 As Integer
Private WDias2 As Integer
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
Private XImpoDto As Double
Private WImpoInteres As Double
Private WDescuento As Double
Private WTasa As Double
Private WImporte As Double
Private WCodIva As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private dada As String
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
Private WAceptada As String
Private Stk(19, 4) As String
Private Envase(5, 2) As String
Private parcial As String
Private Auxiliar(100, 15) As String
Private RestaPedido(100, 3) As String
Private ClavePedido(100)
Private BajaLote(5, 2) As String
Private xLote(100, 12) As String
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
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovenv As Recordset
Dim spMovenv As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim WSaldo4 As Double
Dim WSaldo5 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim XSaldo4 As String
Dim XSaldo5 As String
Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim Compara As Double
Private WCodIb As Integer
Private WImpoIb As Double
Dim ControlLote(5, 2) As String
Dim WSal As Double
Private WAdicional As Double

Private Sub Calcula_FechaVto()

    spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias1 = rstPago!Dias
        WPlazo1 = rstPago!Plazo
        WTasa = rstPago!Tasa
        WDescuento = rstPago!Descuento
        WPago = rstPago!Nombre
        rstPago.Close
    End If
    
    WFecha = Fecha.Text
    Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
    spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WDias2 = rstPago!Dias
        WPlazo2 = rstPago!Plazo
        rstPago.Close
   End If
    
    Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)

End Sub

Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 3
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = ""
    
    DBGrid1.Col = 6
    DBGrid1.Text = "S"
    
    xLote(WRow, 1) = ""
    xLote(WRow, 2) = ""
    xLote(WRow, 3) = ""
    xLote(WRow, 4) = ""
    xLote(WRow, 5) = ""
    xLote(WRow, 6) = ""
    xLote(WRow, 7) = ""
    xLote(WRow, 8) = ""
    xLote(WRow, 9) = ""
    xLote(WRow, 10) = ""
    
End Sub

Private Sub Calcula_Click()

    WNeto = 0
    
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 3
            Precio = DBGrid1.Text
            
            DBGrid1.Col = 4
            Cantidad = DBGrid1.Text
                    
            If Val(Cantidad) <> 0 Then
                WNeto = WNeto + (Val(Cantidad) * Val(Precio))
            End If
                    
        Next iRow
            
    Next A
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0

    Rem If Val(Paridad.Text) <> 0 Then
    Rem    WNeto = WNeto * Val(Paridad.Text)
    Rem End If
    
    XNeto = WNeto
    
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
    
    Select Case WCodIb
        Case 0, 1
            Select Case Val(WCodIva)
                Case 1
                    WImpoIb = WNeto * 0.02
                Case 2, 4, 5, 6
                    WImpoIb = WNeto * 0.025
                Case Else
                    WImpoIb = 0
            End Select
            Call Redondeo(WImpoIb)
        Case Else
            WImpoIb = 0
    End Select
    Compara = WNeto
    Call Redondeo(Compara)
    If Compara < 100 Then
        WImpoIb = 0
    End If
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.21
            WIva2 = WNeto * 0.105
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Case 4
            WIva1 = 0
            WIva2 = 0
        Case Else
            WIva1 = WNeto * 0.21
            Call Redondeo(WIva1)
    End Select
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Neto.Caption = "0.00"
    End If
    
    If WImpoIb <> 0 Then
        Call Convierte1_datos(Str$(WImpoIb), Auxi)
        ImpoIb.Caption = Pusing("###,###.##", Auxi)
            Else
        ImpoIb.Caption = "0.00"
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
    
    WTotal = WNeto + WIva1 + WIva2 + WImpoIb
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    PrgFactup.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    If Val(WEmpresa) = 1 Then
        OPEN_FILE_Ctacte8
        OPEN_FILE_Numero8
        OPEN_FILE_Esta8
    End If
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    Call Verifica_Lote
    If WEstado = "N" Then
        Call Limpia_Click
        Exit Sub
    End If
    
    If Tipoventa.ListIndex = 1 Then
    
        spConsig = "ListaConsig " + "'" + Remito.Text + "'"
        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
        If rstConsig.RecordCount = 0 Then
            m$ = "No Existe el Remito de mercaderia en Consignacion Especificado"
            A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
            Exit Sub
                Else
            If Cliente.Text <> rstConsig!Cliente Then
                m$ = "No coincide el cliente informado con el especificado en el remito"
                A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                Exit Sub
            End If
            rstConsig.Close
        End If
        
        WRenglon = 0
        DBGrid1.Refresh
        
        For A = 0 To 3
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WBase = Val(Right$(Articulo, 3))
                WBaseDy = Val(Left$(Articulo, 2))
                Rem If WBase <= 5 And WBaseDy = "PT" Then
                Rem     Articulo = Left$(Articulo, 7) + "100"
                Rem End If
                
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                    XParam = "'" + Remito.Text + "','" _
                            + Articulo + "'"
                    spConsig = "ListaConsigFactura " + XParam
                    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConsig.RecordCount > 0 Then
                        WSaldo = rstConsig!Cantidad - rstConsig!Facturado
                        If Cantidad > WSaldo Then
                            m$ = "Cantidad insuficiente en consignacion Articulo " + Articulo + " Saldo : " + Str$(WSaldo)
                            A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                            Exit Sub
                        End If
                        rstConsig.Close
                            Else
                        m$ = "No existe este producto en consignacion Articulo " + Articulo
                        A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                        Exit Sub
                    End If
                End If
                                        
            Next iRow
        Next A
    End If
    
    If Tipoventa.ListIndex = 0 Then
    
        WRenglon = 0
        DBGrid1.Refresh
        
        For A = 0 To 3
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            For iRow = 0 To 9
            
                WRenglon = WRenglon + 1
                WRow = iRow
                DBGrid1.Row = WRow
                
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                    DBGrid1.Col = 6
                    Rem If DBGrid1.Text <> "S" Then
                    Rem     m$ = "No asigno las partidas a todos los productos"
                    Rem     A% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    Rem     DBGrid1.Refresh
                    Rem     Exit Sub
                    Rem End If
                End If
                
            Next iRow
        Next A
    End If
    
        Call Calcula_Click

        Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
        Rem     WImporte = WNeto
        Rem     WNeto = WNeto / 1.21
        Rem     Call Redondeo(WNeto)
        Rem     WIva1 = WImporte - WNeto
        Rem     WIva2 = 0
        Rem End If
        
        WTipo = "01"
        WNumero = Numero.Text
        WRenglon = "01"
        WCliente = Cliente.Text
        WFecha = Fecha.Text
        WEstado = "0"
        Rem Wvencimiento = Wvencimiento
        Rem WVencimiento1 = WVencimiento1
        Call Convierte_datos(Str$(Total), Auxi)
        XTotalUs = Str$(WTotal / Val(Paridad.Text))
        XTotal = Str$(WTotal)
        XSaldoUs = Str$(WTotal / Val(Paridad.Text))
        XSaldo = Str$(WTotal)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
        WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
        WImpre = "FC"
        XNet = Str$(WNeto)
        XIva1 = Str$(WIva1)
        XIva2 = Str$(WIva2)
        XImpoIb = Str$(WImpoIb)
        XSeguro = ""
        XFlete = ""
        WPedido = Pedido.Text
        WRemito = Remito.Text
        WOrden = Orden.Text
        WParidad = Paridad.Text
        WProvincia = WProv
        XVendedor = Str$(WVendedor)
        XRubro = Str$(WRubro)
        WComprobante = ""
        WAceptada = Str$(Tipoventa.ListIndex)
        Call Ceros(WAceptada, 1)
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
        WClave = "01" + Auxi + "01"
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
                    + XEmpresa + "','" _
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
            
            With rstCtacte8
                .Index = "Clave"
                .AddNew
                !Tipo = "01"
                !Numero = WNumero8
                !Renglon = "00"
                !Cliente = Cliente.Text
                !Fecha = Fecha.Text
                !Estado = "0"
                !Vencimiento = "  /  /    "
                !Vencimiento1 = "  /  /    "
                Call Convierte_datos(Str$(Total), Auxi)
                !Total = (WNeto * WAdicional)
                !Totalus = (WNeto * WAdicional) / Val(Paridad.Text)
                !Saldo = (WNeto * WAdicional)
                !Saldous = (WNeto * WAdicional) / Val(Paridad.Text)
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = "00000000"
                !OrdVencimiento1 = "00000000"
                !Impre = "FC"
                !Neto = (WNeto * WAdicional)
                !Iva1 = 0
                !Iva2 = 0
                !Pedido = 0
                !Remito = 0
                !Orden = ""
                !Paridad = Val(Paridad.Text)
                !Provincia = WProv
                !Vendedor = WVendedor
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
                !Clave = "01" + Auxi + "00"
                !WDate = Date$
                !TipoDescarga = 1
                .Update
            End With
            
        End If
        
        Erase Auxiliar
        Erase RestaPedido
        Auxi = 0
        
        Suma = 0
        Renglon = 0
        Renglon1 = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For A = 0 To 3
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WTipoProDy = Left$(Articulo, 2)
                Rem WBase = Val(Right$(Articulo, 3))
                Rem If WBase <= 5 Then
                Rem     Articulo = Left$(Articulo, 7) + "100"
                Rem End If
                    
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
                
                Rem If WDescuento <> 0 Then
                Rem     XImpoDto = Precio * WDescuento / 100
                Rem     Call Redondeo(XImpoDto)
                Rem     Precio = Precio - XImpoDto
                Rem End If
                    
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 5
                RestaCantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    If WTipoProDy <> "DY" Then
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                            Else
                        WLinea = 16
                    End If
                        
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                            
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                    WTipo = "01"
                    WNumero = Numero.Text
                    XRenglon = Str$(Renglon)
                    WArticulo = Articulo
                    XXCantidad = Str$(Cantidad)
                    XPrecioUs = Str$(Precio)
                    XPrecio = Str$(Precio * Val(Paridad.Text))
                    XImporteUs = Str$(Precio * Cantidad)
                    XImporte = Str$(Precio * Val(Paridad.Text) * Cantidad)
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = ""
                    XCosto1 = ""
                    WCoeficiente = ""
                    WPedido = Pedido.Text
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = Left$(Articulo, 8)
                    If Tipoventa.ListIndex = 1 Then
                        WRemito = "C" + Remito.Text
                            Else
                        WRemito = Remito.Text
                    End If
                    WClave = "01" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    XMarca = ""
                    Rem caca
                    WLote1 = xLote(Suma, 1)
                    WLote2 = xLote(Suma, 3)
                    Wlote3 = xLote(Suma, 5)
                    WLote4 = xLote(Suma, 7)
                    WLote5 = xLote(Suma, 9)
                    WImpo = xLote(Suma, 2)
                    WCanti1 = Str$(WImpo)
                    WImpo = xLote(Suma, 4)
                    WCanti2 = Str$(WImpo)
                    WImpo = xLote(Suma, 6)
                    WCanti3 = Str$(WImpo)
                    WImpo = xLote(Suma, 8)
                    WCanti4 = Str$(WImpo)
                    WImpo = xLote(Suma, 10)
                    WCanti5 = WImpo
                    If WCliente = "G00007" And Left$(WArticulo, 8) = "PT-07581" Then
                        XLinea = "18"
                    End If
                    If WCliente = "G00065" And Left$(WArticulo, 8) = "PT-07581" Then
                        XLinea = "18"
                    End If
                    If WTipoProDy = "DY" Then
                        XTipoproDy = "M"
                        XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                            Else
                        XTipoproDy = "T"
                        XArticuloDy = "  -   -   "
                    End If
                    XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + XRenglon + "','" + WArticulo + "','" _
                                + XXCantidad + "','" + XPrecio + "','" _
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
                                + XImpo + "','" _
                                + XImpoUs + "','" _
                                + XMarca + "','" _
                                + WLote1 + "','" + WCanti1 + "','" _
                                + WLote2 + "','" + WCanti2 + "','" _
                                + Wlote3 + "','" + WCanti3 + "','" _
                                + WLote4 + "','" + WCanti4 + "','" _
                                + WLote5 + "','" + WCanti5 + "','" _
                                + XTipoproDy + "','" + XArticuloDy + "'"
                    
                    spEstadistica = "AltaEstadistica " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If WAdicional > 0 Then
                        Auxi1 = Str$(WNumero8)
                        Call Ceros(Auxi1, 8)
                        With rstEsta8
                            .Index = "Clave"
                            .AddNew
                            !Tipo = "01"
                            !Numero = WNumero8
                            !Renglon = Renglon
                            !Articulo = Articulo
                            !Cantidad = Cantidad
                            !Precio = Precio * WAdicional
                            !PrecioUs = Precio * WAdicional / Val(Paridad.Text)
                            !Importe = Precio * Cantidad * WAdicional
                            !ImporteUs = Precio * Cantidad * WAdicional / Val(Paridad.Text)
                            !Cliente = Cliente.Text
                            !Paridad = Val(Paridad.Text)
                            !Vendedor = WVendedor
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
                            !Clave = "01" + Auxi1 + Auxi
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
                    Auxiliar(Renglon, 5) = WLote1
                    Auxiliar(Renglon, 6) = WCanti1
                    Auxiliar(Renglon, 7) = WLote2
                    Auxiliar(Renglon, 8) = WCanti2
                    Auxiliar(Renglon, 9) = Wlote3
                    Auxiliar(Renglon, 10) = WCanti3
                    Auxiliar(Renglon, 11) = WLote4
                    Auxiliar(Renglon, 12) = WCanti4
                    Auxiliar(Renglon, 13) = WLote5
                    Auxiliar(Renglon, 14) = WCanti5
                    Auxiliar(Renglon, 15) = RestaCantidad
                        
                End If
                
                If RestaCantidad <> 0 Then
                    Renglon1 = Renglon1 + 1
                    RestaPedido(Renglon1, 1) = Articulo
                    RestaPedido(Renglon1, 2) = RestaCantidad
                    RestaPedido(Renglon1, 3) = ClavePedido(WRenglon)
                End If
                                        
            Next iRow
            
        Next A
        
        Erase Auxiliar
        Erase RestaPedido
        Auxi = 0
        
        Suma = 0
        Renglon = 0
        Renglon1 = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        For A = 0 To 3
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                Suma = Suma + 1
                WRenglon = WRenglon + 1
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Articulo = DBGrid1.Text
                WTipoProDy = Left$(Articulo, 2)
                Rem WBase = Val(Right$(Articulo, 3))
                Rem If WBase <= 5 Then
                Rem     Articulo = Left$(Articulo, 7) + "100"
                Rem End If
                    
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
                
                Rem If WDescuento <> 0 Then
                Rem     XImpoDto = Precio * WDescuento / 100
                Rem     Call Redondeo(XImpoDto)
                Rem     Precio = Precio - XImpoDto
                Rem End If
                    
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                DBGrid1.Col = 5
                RestaCantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    If WTipoProDy <> "DY" Then
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                            Else
                        WLinea = 16
                    End If
                        
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                            
                    Auxi1 = Str$(Numero.Text)
                    Call Ceros(Auxi1, 8)
                    WTipo = "01"
                    WNumero = Numero.Text
                    XRenglon = Str$(Renglon)
                    WArticulo = Articulo
                    XXCantidad = Str$(Cantidad)
                    XPrecioUs = Str$(Precio / Val(Paridad.Text))
                    XPrecio = Str$(Precio)
                    XImporteUs = Str$((Precio * Cantidad) / Val(Paridad.Text))
                    XImporte = Str$(Precio * Cantidad)
                    WCliente = Cliente.Text
                    WParidad = Paridad.Text
                    XVendedor = Str$(WVendedor)
                    XRubro = Str$(WRubro)
                    XLinea = Str$(WLinea)
                    XCosto2 = ""
                    XCosto1 = ""
                    WCoeficiente = ""
                    WPedido = Pedido.Text
                    WFecha = Fecha.Text
                    WImporte1 = ""
                    WImporte2 = ""
                    WImporte3 = ""
                    WImporte4 = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XArticulo = Left$(Articulo, 8)
                    If Tipoventa.ListIndex = 1 Then
                        WRemito = "C" + Remito.Text
                            Else
                        WRemito = Remito.Text
                    End If
                    WClave = "01" + Auxi1 + Auxi
                    WDate = Date$
                    XCanti = ""
                    XImpo = ""
                    XImpoUs = ""
                    XMarca = ""
                    Rem caca
                    WLote1 = xLote(Suma, 1)
                    WLote2 = xLote(Suma, 3)
                    Wlote3 = xLote(Suma, 5)
                    WLote4 = xLote(Suma, 7)
                    WLote5 = xLote(Suma, 9)
                    WImpo = xLote(Suma, 2)
                    WCanti1 = Str$(WImpo)
                    WImpo = xLote(Suma, 4)
                    WCanti2 = Str$(WImpo)
                    WImpo = xLote(Suma, 6)
                    WCanti3 = Str$(WImpo)
                    WImpo = xLote(Suma, 8)
                    WCanti4 = Str$(WImpo)
                    WImpo = xLote(Suma, 10)
                    WCanti5 = WImpo
                    If WCliente = "G00007" And WArticulo = "PT-07581-100" Then
                        XLinea = "18"
                    End If
                    If WCliente = "G00065" And WArticulo = "PT-07581-100" Then
                        XLinea = "18"
                    End If
                    If WTipoProDy = "DY" Then
                        XTipoproDy = "M"
                        XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                            Else
                        XTipoproDy = "T"
                        XArticuloDy = "  -   -   "
                    End If
                    XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + XRenglon + "','" + WArticulo + "','" _
                                + XXCantidad + "','" + XPrecio + "','" _
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
                                + XImpo + "','" _
                                + XImpoUs + "','" _
                                + XMarca + "','" _
                                + WLote1 + "','" + WCanti1 + "','" _
                                + WLote2 + "','" + WCanti2 + "','" _
                                + Wlote3 + "','" + WCanti3 + "','" _
                                + WLote4 + "','" + WCanti4 + "','" _
                                + WLote5 + "','" + WCanti5 + "','" _
                                + XTipoproDy + "','" + XArticuloDy + "'"
                    
                    spEstadistica = "AltaEstadistica " + XParam
                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                    If WAdicional > 0 Then
                        Auxi1 = Str$(WNumero8)
                        Call Ceros(Auxi1, 8)
                        With rstEsta8
                            .Index = "Clave"
                            .AddNew
                            !Tipo = "01"
                            !Numero = WNumero8
                            !Renglon = Renglon
                            !Articulo = Articulo
                            !Cantidad = Cantidad
                            !Precio = Precio * Val(Paridad.Text) * WAdicional
                            !PrecioUs = Precio * WAdicional
                            !Importe = Precio * Cantidad * Val(Paridad.Text) * WAdicional
                            !ImporteUs = Precio * Cantidad * WAdicional
                            !Cliente = Cliente.Text
                            !Paridad = Val(Paridad.Text)
                            !Vendedor = WVendedor
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
                            !Clave = "01" + Auxi1 + Auxi
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
                    Auxiliar(Renglon, 5) = WLote1
                    Auxiliar(Renglon, 6) = WCanti1
                    Auxiliar(Renglon, 7) = WLote2
                    Auxiliar(Renglon, 8) = WCanti2
                    Auxiliar(Renglon, 9) = Wlote3
                    Auxiliar(Renglon, 10) = WCanti3
                    Auxiliar(Renglon, 11) = WLote4
                    Auxiliar(Renglon, 12) = WCanti4
                    Auxiliar(Renglon, 13) = WLote5
                    Auxiliar(Renglon, 14) = WCanti5
                    Auxiliar(Renglon, 15) = RestaCantidad
                        
                End If
                    
                    
                    
                    
                    
                    
                    Auxiliar(Renglon, 1) = Articulo
                    Auxiliar(Renglon, 2) = Cantidad
                    Auxiliar(Renglon, 3) = Precio
                    Auxiliar(Renglon, 4) = WRenglon
                    Auxiliar(Renglon, 5) = WLote1
                    Auxiliar(Renglon, 6) = WCanti1
                    Auxiliar(Renglon, 7) = WLote2
                    Auxiliar(Renglon, 8) = WCanti2
                    Auxiliar(Renglon, 9) = Wlote3
                    Auxiliar(Renglon, 10) = WCanti3
                    Auxiliar(Renglon, 11) = WLote4
                    Auxiliar(Renglon, 12) = WCanti4
                    Auxiliar(Renglon, 13) = WLote5
                    Auxiliar(Renglon, 14) = WCanti5
                    Auxiliar(Renglon, 15) = RestaCantidad
                        
                End If
                
                If RestaCantidad <> 0 Then
                    Renglon1 = Renglon1 + 1
                    RestaPedido(Renglon1, 1) = Articulo
                    RestaPedido(Renglon1, 2) = RestaCantidad
                    RestaPedido(Renglon1, 3) = ClavePedido(WRenglon)
                End If
                                        
            Next iRow
            
        Next A
        
        For da = 1 To Renglon
        
            Articulo = Auxiliar(da, 1)
            Cantidad = Auxiliar(da, 2)
            Precio = Auxiliar(da, 3)
            WRenglon = Auxiliar(da, 4)
            WLote1 = Auxiliar(da, 5)
            WCanti1 = Auxiliar(da, 6)
            WLote2 = Auxiliar(da, 7)
            WCanti2 = Auxiliar(da, 8)
            Wlote3 = Auxiliar(da, 9)
            WCanti3 = Auxiliar(da, 10)
            WLote4 = Auxiliar(da, 11)
            WCanti4 = Auxiliar(da, 12)
            WLote5 = Auxiliar(da, 13)
            WCanti5 = Auxiliar(da, 14)
            RestaCantidad = Auxiliar(da, 15)
            WTipoProDy = Left$(Articulo, 2)
            If WTipoProDy = "DY" Then
                XTipoproDy = "M"
                XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                    Else
                XTipoproDy = "T"
                XArticuloDy = "  -   -   "
            End If
            
            Select Case Tipoventa.ListIndex
                Case 1
                    XParam = "'" + Remito.Text + "','" _
                            + Articulo + "'"
                    spConsig = "ListaConsigFactura " + XParam
                    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConsig.RecordCount > 0 Then
                        WClave = rstConsig!Clave
                        WFacturado = Str$(rstConsig!Facturado + Cantidad)
                        rstConsig.Close
                
                        XParam = "'" + WClave + "','" _
                                + WFacturado + "'"
                                           
                        spConsig = "ModificaConsigFacturado " + XParam
                        Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
                    End If
            
                Case Else
                    If XTipoproDy = "M" Then
                    
                        spArticulo = "ConsultaArticulo " + "'" + XArticuloDy + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WCodigo = XArticuloDy
                            WPedido = Str$(rstArticulo!Venta - RestaCantidad)
                            WSalidas = Str$(rstArticulo!Salidas + Cantidad)
                            WDate = Date$
                            rstArticulo.Close
                            XParam = "'" + WCodigo + "','" _
                                    + WPedido + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                            spArticulo = "ModificaArticuloFacturas " + XParam
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        BajaLote(1, 1) = WLote1
                        BajaLote(1, 2) = WCanti1
                        BajaLote(2, 1) = WLote2
                        BajaLote(2, 2) = WCanti2
                        BajaLote(3, 1) = Wlote3
                        BajaLote(3, 2) = WCanti3
                        BajaLote(4, 1) = WLote4
                        BajaLote(4, 2) = WCanti4
                        BajaLote(5, 1) = WLote5
                        BajaLote(5, 2) = WCanti5
                        
                        For XDa = 1 To 5
                        
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                            
                            If Val(lote1) <> 0 Then
                        
                                XParam = "'" + lote1 + "','" _
                                        + XArticuloDy + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo - Val(Cantidad1))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + WClave + "','" _
                                                 + WDate + "','" _
                                                 + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            
                                            Else
                                
                                    XParam = "'" + XArticuloDy + "','" _
                                                 + lote1 + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad1))
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
                        Next XDa
                    
                            Else
                            
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WCodigo = Articulo
                            WPedido = Str$(rstTerminado!Pedido - RestaCantidad)
                            WSalidas = Str$(rstTerminado!Salidas + Cantidad)
                            WDate = Date$
                        
                            WLinea = rstTerminado!Linea
                            rstTerminado.Close
                    
                            XParam = "'" + WCodigo + "','" _
                                         + WPedido + "','" _
                                         + WSalidas + "','" _
                                         + WDate + "'"
                                            
                            spTerminado = "ModificaTerminadoFacturas " + XParam
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        BajaLote(1, 1) = WLote1
                        BajaLote(1, 2) = WCanti1
                        BajaLote(2, 1) = WLote2
                        BajaLote(2, 2) = WCanti2
                        BajaLote(3, 1) = Wlote3
                        BajaLote(3, 2) = WCanti3
                        BajaLote(4, 1) = WLote4
                        BajaLote(4, 2) = WCanti4
                        BajaLote(5, 1) = WLote5
                        BajaLote(5, 2) = WCanti5
                        
                        For XDa = 1 To 5
                    
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                
                            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                rstTerminado.Close
                            End If
                            
                            lote1 = BajaLote(XDa, 1)
                            Cantidad1 = BajaLote(XDa, 2)
                            
                            If WControla = 0 And Val(lote1) <> 0 Then
                                XParam = "'" + lote1 + "','" _
                                             + Articulo + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                
                                    WClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo - Val(Cantidad1))
                                    WDate = Date$
                                    rstHoja.Close
                                    
                                    XParam = "'" + WClave + "','" _
                                                 + WDate + "','" _
                                                 + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                
                                        Else
                                    
                                    XParam = "'" + Articulo + "','" _
                                                 + lote1 + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo - Val(Cantidad1))
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
                    
                        Next XDa
                        
                    End If
                    
            End Select
            
            If XTipoproDy = "M" Then
            
                ClavePrecioMp = Cliente.Text + XArticuloDy
            
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePrecioMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
            
                    WFecha1 = ""
                    WFactura1 = ""
                    WPrecio1 = ""
                    WCantidad1 = ""
                
                    WFecha2 = ""
                    WFactura2 = ""
                    WPrecio2 = ""
                    WCantidad2 = ""
                
                    WFecha3 = ""
                    WFactura3 = ""
                    WPrecio3 = ""
                    WCantidad3 = ""
                
                    WFecha4 = ""
                    WFactura4 = ""
                    WPrecio4 = ""
                    WCantidad4 = ""
                
                    WFecha5 = ""
                    WFactura5 = ""
                    WPrecio5 = ""
                    WCantidad5 = ""
                
                    If rstPrecios!Cantidad2 <> O Then
                        WFecha1 = rstPrecios!fecha2
                        WFactura1 = rstPrecios!Factura2
                        WPrecio1 = Str$(rstPrecios!Precio2)
                        WCantidad1 = Str$(rstPrecios!Cantidad2)
                    End If
                                
                    If rstPrecios!Cantidad3 <> O Then
                        WFecha2 = rstPrecios!Fecha3
                        WFactura2 = rstPrecios!Factura3
                        WPrecio2 = Str$(rstPrecios!Precio3)
                        WCantidad2 = Str$(rstPrecios!Cantidad3)
                    End If
                                
                    If rstPrecios!Cantidad4 <> O Then
                        WFecha3 = rstPrecios!Fecha4
                        WFactura3 = rstPrecios!Factura4
                        WPrecio3 = Str$(rstPrecios!Precio4)
                        WCantidad3 = Str$(rstPrecios!Cantidad4)
                    End If
                                
                    If rstPrecios!Cantidad5 <> O Then
                        WFecha4 = rstPrecios!Fecha5
                        WFactura4 = rstPrecios!Factura5
                        WPrecio4 = Str$(rstPrecios!Precio5)
                        WCantidad4 = Str$(rstPrecios!Cantidad5)
                    End If
                                
                    WFecha5 = Fecha.Text
                    WFactura5 = Numero.Text
                    WPrecio5 = Str$(Precio / Val(Paridad.Text))
                    WCantidad5 = Str$(Cantidad)
                                
                    WDate = Date$
                
                    rstPreciosMp.Close
                
                    XParam = "'" + ClavePrecioMp + "','" _
                            + WFecha1 + "','" _
                            + WFactura1 + "','" _
                            + WPrecio1 + "','" _
                            + WCantidad1 + "','" _
                            + WFecha2 + "','" _
                            + WFactura2 + "','" _
                            + WPrecio2 + "','" _
                            + WCantidad2 + "','" _
                            + WFecha3 + "','" _
                            + WFactura3 + "','" _
                            + WPrecio3 + "','" _
                            + WCantidad3 + "','" _
                            + WFecha4 + "','" _
                            + WFactura4 + "','" _
                            + WPrecio4 + "','" _
                            + WCantidad4 + "','" _
                            + WFecha5 + "','" _
                            + WFactura5 + "','" _
                            + WPrecio5 + "','" _
                            + WCantidad5 + "','" _
                            + WDate + "'"
                                           
                    spPreciosMp = "ModificaPreciosFacturaMp " + XParam
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
                    Else
                
                ClavePrecio = Cliente.Text + Articulo
            
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecio + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
            
                    WFecha1 = ""
                    WFactura1 = ""
                    WPrecio1 = ""
                    WCantidad1 = ""
                
                    WFecha2 = ""
                    WFactura2 = ""
                    WPrecio2 = ""
                    WCantidad2 = ""
                
                    WFecha3 = ""
                    WFactura3 = ""
                    WPrecio3 = ""
                    WCantidad3 = ""
                
                    WFecha4 = ""
                    WFactura4 = ""
                    WPrecio4 = ""
                    WCantidad4 = ""
                
                    WFecha5 = ""
                    WFactura5 = ""
                    WPrecio5 = ""
                    WCantidad5 = ""
                
                    If rstPrecios!Cantidad2 <> O Then
                        WFecha1 = rstPrecios!fecha2
                        WFactura1 = rstPrecios!Factura2
                        WPrecio1 = Str$(rstPrecios!Precio2)
                        WCantidad1 = Str$(rstPrecios!Cantidad2)
                    End If
                                
                    If rstPrecios!Cantidad3 <> O Then
                        WFecha2 = rstPrecios!Fecha3
                        WFactura2 = rstPrecios!Factura3
                        WPrecio2 = Str$(rstPrecios!Precio3)
                        WCantidad2 = Str$(rstPrecios!Cantidad3)
                    End If
                                
                    If rstPrecios!Cantidad4 <> O Then
                        WFecha3 = rstPrecios!Fecha4
                        WFactura3 = rstPrecios!Factura4
                        WPrecio3 = Str$(rstPrecios!Precio4)
                        WCantidad3 = Str$(rstPrecios!Cantidad4)
                    End If
                                
                    If rstPrecios!Cantidad5 <> O Then
                        WFecha4 = rstPrecios!Fecha5
                        WFactura4 = rstPrecios!Factura5
                        WPrecio4 = Str$(rstPrecios!Precio5)
                        WCantidad4 = Str$(rstPrecios!Cantidad5)
                    End If
                                
                    WFecha5 = Fecha.Text
                    WFactura5 = Numero.Text
                    WPrecio5 = Str$(Precio)
                    WCantidad5 = Str$(Cantidad)
                                
                    WDate = Date$
                
                    rstPrecios.Close
                
                    XParam = "'" + ClavePrecio + "','" _
                            + WFecha1 + "','" _
                            + WFactura1 + "','" _
                            + WPrecio1 + "','" _
                            + WCantidad1 + "','" _
                            + WFecha2 + "','" _
                            + WFactura2 + "','" _
                            + WPrecio2 + "','" _
                            + WCantidad2 + "','" _
                            + WFecha3 + "','" _
                            + WFactura3 + "','" _
                            + WPrecio3 + "','" _
                            + WCantidad3 + "','" _
                            + WFecha4 + "','" _
                            + WFactura4 + "','" _
                            + WPrecio4 + "','" _
                            + WCantidad4 + "','" _
                            + WFecha5 + "','" _
                            + WFactura5 + "','" _
                            + WPrecio5 + "','" _
                            + WCantidad5 + "','" _
                            + WDate + "'"
                                           
                    spPrecios = "ModificaPreciosFactura " + XParam
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            End If
        Next da
        
        For da = 1 To Renglon1
        
            Articulo = RestaPedido(da, 1)
            Cantidad = RestaPedido(da, 2)
            WClavePedido = RestaPedido(da, 3)
            
            XParam = "'" + Left$(WClavePedido, 6) + "','" _
                        + Right$(WClavePedido, 2) + "'"
            spPedido = "ConsultaPedido2 " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                WFacturado = Str$(rstPedido!Facturado + Cantidad)
                If Val(WFacturado) > rstPedido!Cantidad Then
                    WFacturado = Str$(rstPedido!Cantidad)
                End If
                WClavePedido = rstPedido!Clave
                rstPedido.Close
                XParam = "'" + WClavePedido + "','" _
                            + WFacturado + "'"
                                           
                spPedido = "ModificaPedidoFacturas " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        Next da
        
        BajaImpre = "N"
        
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                
                        WTerminado = !Terminado
                        XCodigo = Val(Mid$(WTer3minado, 4, 5))
                        Canti = !Cantidad - !Facturado
                        
                        If Canti > 0 Then
                            BajaImpre = "S"
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
        End If
        
        If BajaImpre = "S" Then
        
            spPedido = "ModificaPedidoVersion " + "'" + Pedido.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            If Left$(WTerminado, 2) = "DY" Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 0 And XCodigo <= 999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 11000 And XCodigo <= 11999 Then
                        XTipoPro = "CO"
                            Else
                        If XCodigo >= 25000 And XCodigo <= 25999 Then
                            XTipoPro = "FA"
                                Else
                            If XCodigo >= 2300 And XCodigo <= 2399 Then
                                XTipoPro = "BI"
                                    Else
                                XTipoPro = "PT"
                            End If
                        End If
                    End If
                End If
            End If
                    
            WLinea = 0
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WLinea = rstTerminado!Linea
                rstTerminado.Close
            End If
            If WLinea = 8 Then
                XTipoPro = "PG"
            End If
            
            Select Case XTipoPro
                Case "CO"
                    XParam = "'" + Pedido.Text + "','" _
                                + "1" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                Case "FA"
                    XParam = "'" + Pedido.Text + "','" _
                                + "4" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                Case "BI"
                    XParam = "'" + Pedido.Text + "','" _
                                + "3" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                Case "PT"
                    XParam = "'" + Pedido.Text + "','" _
                                + "2" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                Case "PG"
                    XParam = "'" + Pedido.Text + "','" _
                                + "5" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    
                    WMarca = "X"
                    XParam = "'" + Pedido.Text + "','" _
                            + WMarca + "'"
                    spPedido = "ModificaPedidoPigmentos " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                Case Else
                    XParam = "'" + Pedido.Text + "','" _
                                + "0" + "'"
                    spPedido = "ModificaPedidoTipoPedido " + XParam
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            End Select
            
        End If
        
        spNumero = "ConsultaNumero " + "'" + "01" + "'"
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        If rstNumero.RecordCount > 0 Then
            WCodigo = "01"
            WNumero = Numero.Text
            rstNumero.Close
            XParam = "'" + WCodigo + "','" _
                         + WNumero + "'"
            spNumero = "ModificaNumero " + XParam
            Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
        Rem Listado.Destination = 1
        Rem Listado.Action = 1
        
        Call Calcula_Saldo
        
        Erase Envase
        Envase(1, 1) = Envase1.Text
        Envase(2, 1) = Envase2.Text
        Envase(3, 1) = Envase3.Text
        Envase(4, 1) = Envase4.Text
        Envase(5, 1) = Envase5.Text
        
        Envase(1, 2) = Canti1.Text
        Envase(2, 2) = Canti2.Text
        Envase(3, 2) = Canti3.Text
        Envase(4, 2) = Canti4.Text
        Envase(5, 2) = Canti5.Text
        
        For XDa = 1 To 5
            For da = 1 To 9
                If Val(Envase(XDa, 1)) = Val(Stk(da, 1)) Then
                    Stk(da, 3) = Val(Envase(XDa, 2))
                End If
            Next da
        Next XDa
        
        For da = 1 To 9
            Stk(da, 4) = Str$(Val(Stk(da, 2)) + Val(Stk(da, 3)))
        Next da
        
        Renglon = 0
        
        For da = 1 To 5
        
            If Val(Envase(da, 2)) <> 0 Then
            
                Renglon = Renglon + 1
                    
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Remito.Text))
                Call Ceros(Auxi1, 6)
                    
                WTipo = "1"
                WCodigo = Str$(Val(Remito.Text) + 100000)
                WRenglon = Str$(Renglon)
                WCliente = Cliente.Text
                WFecha = Fecha.Text
                WEnvase = Envase(da, 1)
                WCantidad = Envase(da, 2)
                WMovimiento = "S"
                WFechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                WClave = Auxi1 + Auxi
                
                XParam = "'" + WClave + "','" _
                        + WTipo + "','" _
                        + WCodigo + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WFechaord + "','" _
                        + WCliente + "','" _
                        + WEnvase + "','" _
                        + WMovimiento + "','" _
                        + WCantidad + "'"
                    
                spMovenv = "AltaMovenv " + XParam
                Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next da
        
        Call Impresion
        If Tipoventa.ListIndex <> 1 Then
            Call Impresion_Remito
        End If
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
    Exit Sub

WError:
     Resume Next
        
End Sub

Private Sub Limpia_Click()

    CargaLote.Visible = False
    Erase xLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    WAdicional = 0
    
    For A = 0 To 3
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
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
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
    
    Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 7
                Rem Select Case KeyCode
                Rem     Case 13
                Rem         DBGrid1.Col = 4
                Rem         DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                Rem         DBGrid1.Col = 5
                Rem         DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                Rem         DBGrid1.Col = 0
                Rem         XTerminado = DBGrid1.Text
                Rem         DBGrid1.Col = 4
                Rem         XCantidad = Val(DBGrid1.Text)
                Rem         WRow = DBGrid1.Row
                Rem
                Rem         Rem If DBGrid1.Row < 40 Then
                Rem         Rem    DBGrid1.Row = DBGrid1.Row + 1
                Rem         Rem    WRow = DBGrid1.Row
                Rem         Rem    DBGrid1.Col = 4
                Rem         Rem    KeyCode = 0
                Rem         Rem End If
                Rem         Rem Call Calcula_Click
                Rem         Rem DBGrid1.Row = WRow
                Rem
                Rem         If Tipoventa.ListIndex = 0 Then
                Rem             CargaLote.Visible = True
                Rem             WLote1.Text = ""
                Rem             WCanti1.Text = ""
                Rem             WLote2.Text = ""
                Rem             WCanti2.Text = ""
                Rem             Wlote3.Text = ""
                Rem             WCanti3.Text = ""
                Rem             WLote4.Text = ""
                Rem             WCanti4.Text = ""
                Rem             WLote5.Text = ""
                Rem             WCanti5.Text = ""
                Rem
                Rem             If Val(xLote(WRow, 1)) <> 0 Then
                Rem                 WLote1.Text = xLote(WRow, 1)
                Rem                 WCanti1.Text = xLote(WRow, 2)
                Rem             End If
                Rem             If Val(xLote(WRow, 3)) <> 0 Then
                Rem                 WLote2.Text = xLote(WRow, 3)
                Rem                 WCanti2.Text = xLote(WRow, 4)
                Rem             End If
                Rem             If Val(xLote(WRow, 5)) <> 0 Then
                Rem                 Wlote3.Text = xLote(WRow, 5)
                Rem                 WCanti3.Text = xLote(WRow, 6)
                Rem             End If
                Rem             If Val(xLote(WRow, 7)) <> 0 Then
                Rem                 WLote4.Text = xLote(WRow, 7)
                Rem                 WCanti4.Text = xLote(WRow, 6)
                Rem             End If
                Rem             If Val(xLote(WRow, 9)) <> 0 Then
                Rem                 WLote5.Text = xLote(WRow, 9)
                Rem                 WCanti5.Text = xLote(WRow, 10)
                Rem             End If
                Rem
                Rem             WLote1.SetFocus
                Rem         End If
                Rem
                Rem     Case Else
                Rem         Rem If KeyCode <> 0 Then Stop
                Rem
                Rem End Select
            
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
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    Tipoventa.Clear
    
    Tipoventa.AddItem "Venta Normal"
    Tipoventa.AddItem "Mercaderia en Consignacion"
    
    Tipoventa.ListIndex = 0

    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 6, 0 To 40)

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
For i = 0 To 6
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1400
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cant. Entregar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Cant. Descontar"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 6
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

    DBGrid1.Font.Bold = True

    Erase xLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    Wlote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    WAdicional = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "01" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Numero.SetFocus
     
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
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
    WNeto = 0
    XEntra = "S"
    
    Erase Auxiliar
    Erase ClavePedido
    Erase xLote
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad - !Facturado
                    
                    If Canti <> 0 Then
                    
                        XCanti1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                        XCanti2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                        
                        If XCanti1 <> 0 Or XCanti2 <> 0 Then
                
                            Renglon = Renglon + 1
                
                            Lugar1 = Int((Renglon - 1) / 10) * 10
                            Lugar2 = Renglon - Lugar1
                
                            DBGrid1.FirstRow = Lugar1
                            DBGrid1.Row = Lugar2 - 1
                
                            DBGrid1.Col = 0
                            DBGrid1.Text = !Terminado
                            Auxi1 = !Terminado
                
                            DBGrid1.Col = 2
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                
                            DBGrid1.Col = 3
                            DBGrid1.Text = Pusing("###,###.##", Str$(!Precio * Val(Paridad.Text)))
                
                            XCantidad1 = IIf(IsNull(!Cantidad1), "0", !Cantidad1)
                            DBGrid1.Col = 4
                            DBGrid1.Text = Pusing("###,###.##", Str$(XCantidad1))
                    
                            xCantidad2 = IIf(IsNull(!Cantidad2), "0", !Cantidad2)
                            DBGrid1.Col = 5
                            DBGrid1.Text = Pusing("###,###.##", Str$(xCantidad2))
                    
                            DBGrid1.Col = 6
                            DBGrid1.Text = "S"
                    
                            
                            If XEntra = "S" Then
                                Envase1.Text = IIf(IsNull(!Env1), "", !Env1)
                                Envase2.Text = IIf(IsNull(!Env2), "", !Env2)
                                Envase3.Text = IIf(IsNull(!Env3), "", !Env3)
                                Envase4.Text = IIf(IsNull(!Env4), "", !Env4)
                                Envase5.Text = IIf(IsNull(!Env5), "", !Env5)
                                Canti1.Text = IIf(IsNull(!CantiEnv1), "", !CantiEnv1)
                                Canti2.Text = IIf(IsNull(!CantiEnv2), "", !CantiEnv2)
                                Canti3.Text = IIf(IsNull(!CantiEnv3), "", !CantiEnv3)
                                Canti4.Text = IIf(IsNull(!CantiEnv4), "", !CantiEnv4)
                                Canti5.Text = IIf(IsNull(!CantiEnv5), "", !CantiEnv5)
                                XEntra = ""
                            End If
                            
                            xLote(Renglon, 1) = IIf(IsNull(!lote1), "", !lote1)
                            xLote(Renglon, 2) = IIf(IsNull(!CantiLote1), "", !CantiLote1)
                            xLote(Renglon, 3) = IIf(IsNull(!lote2), "", !lote2)
                            xLote(Renglon, 4) = IIf(IsNull(!CantiLote2), "", !CantiLote2)
                            xLote(Renglon, 5) = IIf(IsNull(!lote3), "", !lote3)
                            xLote(Renglon, 6) = IIf(IsNull(!CantiLote3), "", !CantiLote3)
                            xLote(Renglon, 7) = IIf(IsNull(!lote4), "", !lote4)
                            xLote(Renglon, 8) = IIf(IsNull(!CantiLote4), "", !CantiLote4)
                            xLote(Renglon, 9) = IIf(IsNull(!lote5), "", !lote5)
                            xLote(Renglon, 10) = IIf(IsNull(!CantiLote4), "", !CantiLote5)
                    
                            Auxiliar(Renglon, 1) = Auxi1
                            Auxiliar(Renglon, 2) = Canti
                            
                            ClavePedido(Renglon) = !Clave
                            
                        End If
                        
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
            Else
        Descri1.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri2.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri3.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri4.Caption = ""
    End If
                        
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
            Else
        Descri5.Caption = ""
    End If
    
    WConpago = 0
    
    WRenglon = Renglon
    Renglon = 0
    
    For da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(da, 1)
        Canti = Auxiliar(da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) = "DY" Then
            WTipoPro = "M"
                Else
            WTipoPro = "T"
        End If
        
        Select Case WTipoPro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                ClavePreciosMp = Cliente.Text + WArti
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                    Precio = rstPreciosMp!Precio * Val(Paridad.Text)
                
                    WConpago = IIf(IsNull(rstPreciosMp!Pago), 0, rstPreciosMp!Pago)
            
                    rstPreciosMp.Close
                End If

                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstArticulo!Descripcion
                    
                    rstArticulo.Close
                End If

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
            
            Case "T"
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
            
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                    Precio = rstPrecios!Precio * Val(Paridad.Text)
                
                    WConpago = IIf(IsNull(rstPrecios!Pago), 0, rstPrecios!Pago)
            
                    rstPrecios.Close
                End If

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
        End Select
        
    Next da
    
    If WConpago <> 0 Then
        WPago1 = WConpago
        WPago2 = WConpago
        
        spPago = "ConsultaPago " + "'" + Str$(WPago1) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDias1 = rstPago!Dias
            WPlazo1 = rstPago!Plazo
            WTasa = rstPago!Tasa
            WDescuento = rstPago!Descuento
            WPago = rstPago!Nombre
            rstPago.Close
        End If
        
        WFecha = Fecha.Text
        Call Calcula_vencimiento(WFecha, WDias1, Wvencimiento)
    
        spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDias2 = rstPago!Dias
            WPlazo2 = rstPago!Plazo
            rstPago.Close
        End If
        
        Call Calcula_vencimiento(WFecha, WDias2, WVencimiento1)
        
    End If
    
    Call Calcula_Click

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
    
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub Proceso1_Click()

    WNeto = 0

    For A = 0 To 3
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
    Erase Auxiliar
    
    
    XParam = "'" + "01" + "','" _
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
                    DBGrid1.Text = rstEstadistica!Articulo
                    Auxi1 = rstEstadistica!Articulo
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Precio)
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", dada)
                
                    dada = Str$(rstEstadistica!Paridad)
                    Paridad.Text = Pusing("###,###.##", dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!Precio)
                    End If
                    
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
    
    For da = 1 To XRenglon
    
        Auxi1 = Auxiliar(da, 1)
        
        If Left$(Auxi1, 2) = "DY" Then
            WTipoPro = "M"
                Else
            WTipoPro = "T"
        End If
        
        Select Case WTipoPro
            Case "M"
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
            
            Case Else
                ClavePrecios = Cliente.Text + Auxi1
        
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
        End Select
    Next da
    
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
    
    Call Calcula_FechaVto
    Call Calcula_Click

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
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0
    
    DBGrid1.SetFocus
    
    Graba.Enabled = False
    Borra.Enabled = False

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Pedido.Text = rstCtacte!Pedido
            Fecha.Text = rstCtacte!Fecha
            Cliente.Text = rstCtacte!Cliente
            Vencimiento.Text = rstCtacte!Vencimiento
            Remito.Text = rstCtacte!Remito
            Orden.Text = rstCtacte!Orden
            Paridad.Text = rstCtacte!Paridad
            rstCtacte.Close
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!Vendedor
                WProv = rstCliente!Provincia
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WCodIb = rstCliente!Ib
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                WDirentrega = rstCliente!DirEntrega
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                rstCliente.Close
            End If
            Call Proceso1_Click
                    Else
            Rem .Index = "Numero"
            Rem .Seek "=", Val(Numero.Text)
            Rem If .NoMatch = False Then
            Rem     m$ = "Comprobante ya existente"
            Rem     A% = MsgBox(m$, 0, "Ingreso de Facturas")
            Rem     Numero.SetFocus
            Rem        Else
            Rem     WNumero = Numero.Text
            Rem    Rem Call Limpia_Click
            Rem    Numero.Text = WNumero
            Rem    Pedido.SetFocus
            Rem End If
            WNumero = Numero.Text
            Rem Call Limpia_Click
            Numero.Text = WNumero
            Fecha.SetFocus
                
        End If
    End If
End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                A% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                Pedido.SetFocus
                    Else
                Cliente.Text = rstPedido!Cliente
                Orden.Text = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                rstPedido.Close
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    WPago1 = rstCliente!Pago1
                    WPago2 = rstCliente!Pago2
                    WVendedor = rstCliente!Vendedor
                    WRubro = rstCliente!Rubro
                    WCodIva = rstCliente!Iva
                    WCodIb = rstCliente!Ib
                    WRazon = rstCliente!Razon
                    WDireccion = rstCliente!Direccion
                    WLocalidad = rstCliente!Localidad
                    WProv = rstCliente!Provincia
                    WPostal = rstCliente!Postal
                    WCuit = rstCliente!Cuit
                    WDirentrega = rstCliente!DirEntrega
                    WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                    rstCliente.Close
                End If
                Call Proceso_Click
                Call Calcula_FechaVto
                Remito.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            spCambios = "ConsultaCambio  " + "'" + Fecha.Text + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                Paridad.Text = Pusing("###,###.##", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
                Rem m$ = "No exsite paridad cargada para esta fecha"
                Rem a% = MsgBox(m$, 0, "Emision de facturas")
                Rem Fecha.SetFocus
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                Pedido.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                A% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            A% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub ReImpre_Click()
    Call Impresion
    Call Impresion_Remito
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
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

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Orden.SetFocus
    End If
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 4
        DBGrid1.Row = 0
        DBGrid1.SetFocus
    End If
End Sub

Sub Impresion()

    If Val(WEmpresa) = 1 Then
        Open "lpt1" For Output As #1
        Rem Open "DADA.TXT" For Output As #1
            Else
        If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
            Open "lpt1" For Output As #1
            Rem Open "DADA.TXT" For Output As #1
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
            Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
                Else
            Open "DADA1.TXT" For Output As #1
        End If
    End If
    
    Rem Width #1, 255

    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "1" + Chr$(72);
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
    
    Paridad = Val(Paridad.Text)
    Impotot = Val(Total.Caption) / Paridad

    For XX% = 1 To 2
    
        If XX% = 1 Then
            Print #1, ""
                Else
            Print #1, ""
        End If

        If Val(WEmpresa) = 1 Then
            Print #1, ""
            Print #1, ""
        End If
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(WEmpresa) = 1 Then
            Print #1, Tab(59); Fecha.Text
                Else
            Print #1, Tab(57); Fecha.Text
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(8); WRazon
        Print #1, Tab(8); WDireccion
        Print #1, Tab(8); Left$(WLocalidad, 33);
        Print #1, Tab(55); Cliente.Text;
        Print #1, Tab(69); Orden.Text
        Print #1, Tab(8); Provincia(Val(WProv)); " ("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(8); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, ""
        Print #1, Tab(5); Left$(WPago, 40); " ";
        Print #1, Vencimiento.Text;
        Print #1, Tab(60); Remito.Text
        Print #1, ""
        Print #1, ""
        Print #1, Tab(76); "$"

        Impre = 0

        For A = 0 To 3
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(Alinea("##,###.##", DBGrid1.Text))
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                    
                If Cantidad <> 0 Then
                
                    Print #1, Tab(1); Alinea("#####.##", Str$(Cantidad));
                    Print #1, " Kg";
                    Print #1, Tab(15); Left$(Descri, 37);
                    parcial = Str$(Precio * Cantidad)
                    
                    Rem If Val(WCodIva) = 1 Or Val(WCodIva) = 2 Then
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem             Else
                    Rem     Precio = Str$(Val(Precio) * 1.21)
                    Rem     Parcial = Str$(Val(Parcial) * 1.21)
                    Rem     Print #1, Tab(57); Alinea("##,###.##", Str$(Precio));
                    Rem     Print #1, Tab(68); Alinea("###,###.##", Str$(Parcial))
                    Rem End If
                    
                    Print #1, Tab(56); " $ "; Alinea("####.##", Str$(Precio));
                    Print #1, Tab(68); Alinea("###,###.##", parcial)
                    
                    Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next A

        For aa = Impre To 22
                Print #1, ""
        Next aa

        Rem M# = Total# / 100
        Rem GoSub 4630
        

        Print #1, Tab(1); "EL IMPORTE DE ESTA FACTURA REPRESENTA U$S ";
        Print #1, Alinea("###,###.##", Str$(Impotot))
        Print #1, Tab(1); "CALCULADOS A UNA PARIDAD DE $ ";
        Print #1, Alinea("##.##", Str$(Paridad))
        Print #1, Tab(1); "Y DEBERA SER CANCELADO A SU VENCIMIENTO EN DOLARES"
        Print #1, Tab(1); "BILLETE  ESTADOUNIDENSES  O  EN  PESOS  AL  CAMBIO"
        Print #1, Tab(1); "OFICIAL  DEL DIA  DE ACREDITACION DE  LOS  VALORES"
        Print #1, Tab(1); "RECIBIDOS."
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        
        Print #1, Tab(65); " $ "; Alinea("###,###.##", Str$(XNeto))

        If Val(Dto.Caption) <> 0 Then
                Print #1, Tab(56); "Dto"; Alinea("##.##", Str$(WDescuento));
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Dto.Caption)
                        Else
                Print #1, ""
        End If

        If Val(Interes.Caption) <> 0 Then
                Print #1, Tab(56); "Interes";
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Interes.Caption)
                                                  Else
                Print #1, ""
        End If

        Print #1, Tab(3); M1;
        Print #1, Tab(65); " $ "; Alinea("###,###.##", Neto.Caption)
        Print #1, Tab(3); M2;
        If Val(Iva1.Caption) <> 0 Then
                Print #1, Tab(61); "21";
                Print #1, Tab(65); " $ "; Alinea("###,###.##", Iva1.Caption)
                        Else
                Print #1, ""
        End If

        Select Case XX
                Case 1
                        Print #1, Tab(10); "ORIGINAL";
                Case 2
                        Print #1, Tab(10); "DUPLICADO";
                Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                Case Else
        End Select

        If Val(ImpoIb.Caption) <> 0 Then
                Print #1, Tab(56); "Ret.I.B.";
                Print #1, Tab(65); " $ "; Alinea("###,###.##", ImpoIb.Caption)
                        Else
                If Val(Iva2.Caption) <> 0 Then
                    Print #1, Tab(60); "10.5";
                    Print #1, Tab(65); " $ "; Alinea("###,###.##", Iva2.Caption)
                        Else
                    Print #1, ""
                End If
        End If

        Print #1, Tab(65); " $ "; Alinea("###,###.##", Total.Caption); Chr$(12)

        Next XX%

        Close #1
        
End Sub

Sub Impresion_Remito()

        If Val(WEmpresa) = 1 Then
            Rem Open "DADA.TXT" For Output As #1
            Open "lpt1" For Output As #1
                Else
            If Val(WEmpresa) <> 9 And Val(WEmpresa) <> 10 Then
                Rem Open "DADA.TXT" For Output As #1
                Open "lpt1" For Output As #1
                Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
                Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70);
                    Else
                Open "DADA.TXT" For Output As #1
            End If
        End If
  
        Rem  #1, 255

        For FF = 1 To 2

        Print #1, Chr$(27) + Chr$(40) + "19U"
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "2" + Chr$(72)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(53); Fecha.Text
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(7); WRazon
        Print #1, Tab(7); Left$(WDireccion, 33)
        Print #1, Tab(7); Left$(WLocalidad, 33);
        Print #1, Tab(44); Pedido.Text;
        Print #1, Tab(57); Cliente.Text;
        Print #1, Tab(68); Orden.Text
        Print #1, Tab(7); Provincia(Val(WProv)); "("; WPostal; ")"
        Print #1, ""
        Print #1, Tab(7); Iva(Val(WCodIva));
        Print #1, Tab(48); WCuit
        Print #1, ""
        Print #1, Tab(30); WDirentrega;
        Print #1, ""
        If FF = 1 Then
            Print #1, Tab(60); "ORIGINAL"
                Else
            Print #1, Tab(60); "DUPLICADO"
        End If
        Print #1, ""
        
        Impre = 0

        For A = 0 To 3
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                Producto = DBGrid1.Text
                
                DBGrid1.Col = 1
                Descri = DBGrid1.Text
                
                DBGrid1.Col = 3
                Precio = Val(DBGrid1.Text)
            
                DBGrid1.Col = 4
                Cantidad = Val(DBGrid1.Text)
                
                If Cantidad <> 0 Then
                        
                        Print #1, Tab(14); Left$(Descri, 40);
                        Print #1, Tab(58); Alinea("#####.##", Str$(Cantidad));
                        Print #1, " Kg";
                        Print #1, Tab(71); "Netos"
                        Impre = Impre + 1
                End If
                    
            Next iRow
            
        Next A
        
        Select Case Val(WEmpresa)
            Case 4, 8
                For aa = Impre To 17
                    Print #1, ""
                Next aa

                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                Print #1, Tab(3); "Pellital S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional"
                Print #1, ""

                For XDa = 1 To 1
                        For da = 1 To 9
                                If Val(Stk(da, 4)) <> 0 Then
                                        
                                        Select Case da
                                                Case 1
                                                        Lugar = 25
                                                Case 2
                                                        Lugar = 36
                                                Case 3
                                                        Lugar = 47
                                                Case 4
                                                        Lugar = 58
                                                Case 5
                                                        Lugar = 69
                                                Case 6
                                                        Lugar = 80
                                                Case 7
                                                        Lugar = 92
                                                Case 8
                                                        Lugar = 104
                                                Case 9
                                                        Lugar = 116
                                                Case Else
                                        End Select
                                                            
                                        If da = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(da, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(da, XDa);
                                        End If
                                    End If
        
                        Next da
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For da = 1 To 9
            
                                If Val(Stk(da, 4)) <> 0 Then
        
                                        Select Case da
                                            Case 1
                                                Lugar = 16
                                            Case 2
                                                Lugar = 23
                                            Case 3
                                                Lugar = 31
                                            Case 4
                                                Lugar = 38
                                            Case 5
                                                Lugar = 45
                                            Case 6
                                                Lugar = 52
                                            Case 7
                                                Lugar = 59
                                            Case 8
                                                Lugar = 66
                                            Case 9
                                                Lugar = 73
                                            Case Else
                                    End Select
        
                                    If Val(Stk(da, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(da, XDa))));
                                    End If
        
                            End If
                    Next da
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, ""
                Print #1, ""
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
            
            Case Else
                For aa = Impre To 19
                    Print #1, ""
                Next aa

                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
                Print #1, Tab(3); "Surfactan S.A. no se responsabiliza por los daños que pudiera causar la aplicación inadecuada de estos productos,"
                Print #1, Tab(3); "el reuso de envases o la mala disposición final de los residuos generados a partir de los mismos."
                Print #1, Tab(3); "Los residuos generados a partir de los productos remitidos con  este comprobante y que presenten riesgos para"
                Print #1, Tab(3); "la salud o para el medio ambiente, deberán ser destruidos y dispuestos según lo establecen las reglamentaciones "
                Print #1, Tab(3); "vigentes del ámbito municipal, provincial y nacional"
                Print #1, ""

                For XDa = 1 To 1
                        For da = 1 To 9
                                If Val(Stk(da, 4)) <> 0 Then
                                        
                                        Select Case da
                                                Case 1
                                                        Lugar = 22
                                                Case 2
                                                        Lugar = 33
                                                Case 3
                                                        Lugar = 44
                                                Case 4
                                                        Lugar = 55
                                                Case 5
                                                        Lugar = 66
                                                Case 6
                                                        Lugar = 77
                                                Case 7
                                                        Lugar = 89
                                                Case 8
                                                        Lugar = 101
                                                Case 9
                                                        Lugar = 113
                                                Case Else
                                        End Select
                                                            
                                        If da = 9 Then
                                            Digi = 10
                                                    Else
                                            Digi = 10
                                        End If
                                
                                        spEnvases = "ConsultaEnvases " + "'" + Str$(Val(Stk(da, XDa))) + "'"
                                        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvases.RecordCount > 0 Then
                                            Print #1, Tab(Lugar); Left$(rstEnvases!Abreviatura, Digi);
                                            rstEnvases.Close
                                                    Else
                                            Print #1, Tab(Lugar); Stk(da, XDa);
                                        End If
                                    End If
        
                        Next da
                        Print #1, ""
        
                Next XDa
        
                Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
                For XDa = 2 To 4
                        For da = 1 To 9
            
                                If Val(Stk(da, 4)) <> 0 Then
        
                                        Select Case da
                                            Case 1
                                                Lugar = 14
                                            Case 2
                                                Lugar = 21
                                            Case 3
                                                Lugar = 29
                                            Case 4
                                                Lugar = 36
                                            Case 5
                                                Lugar = 43
                                            Case 6
                                                Lugar = 50
                                            Case 7
                                                Lugar = 57
                                            Case 8
                                                Lugar = 64
                                            Case 9
                                                Lugar = 71
                                            Case Else
                                    End Select
        
                                    If Val(Stk(da, XDa)) <> 0 Then
                                            Print #1, Tab(Lugar); Alinea("####", Str$(Val(Stk(da, XDa))));
                                    End If
        
                            End If
                    Next da
        
                    Print #1, ""
                    Print #1, ""
                
                Next XDa
        
                Print #1, ""
                Select Case XX
                    Case 1
                        Print #1, Tab(10); "ORIGINAL";
                    Case 2
                        Print #1, Tab(10); "DUPLICADO";
                    Case 3
                        Print #1, Tab(10); "TRIPLICADO";
                    Case Else
                End Select
                Print #1, Tab(10); "Nro. Control : "; Remito.Text
                Print #1, Chr$(12)
                
        End Select

        Next FF

        Close #1

End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
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
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
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
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
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
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
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
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Calcula_Saldo()

    Rem On Error GoTo Error_saldo

    Erase Stk

    Stk(1, 1) = "020"
    Stk(2, 1) = "021"
    Stk(3, 1) = "022"
    Stk(4, 1) = "023"
    Stk(5, 1) = "024"
    Stk(6, 1) = "025"
    Stk(7, 1) = "026"
    Stk(8, 1) = "030"
    Stk(9, 1) = "028"

    XParam = "'" + Cliente.Text + "','" _
                + Cliente.Text + "'"

    spMovenv = "ListaMovenvDesdeHastaCliente " + XParam
    Set rstMovenv = db.OpenRecordset(spMovenv, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovenv.RecordCount > 0 Then
    
        With rstMovenv
            .MoveFirst
            Do
                If .EOF = False Then

                    For da = 1 To 9
                        If Val(Stk(da, 1)) = !Envase Then
                            If !Movimiento = "S" Then
                                Stk(da, 2) = Str$(Val(Stk(da, 2)) + !Cantidad)
                                    Else
                                Stk(da, 2) = Str$(Val(Stk(da, 2)) - !Cantidad)
                            End If
                        End If
                    
                    Next da
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovenv.Close
    End If

End Sub

Private Sub Verifica_Lote()

    Renglon = 0
    Renglon1 = 0
    WRenglon = 0
    DBGrid1.Refresh
        
    For A = 0 To 3
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
            
            Suma = Suma + 1
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            WTipoProDy = Left$(Articulo, 2)
            
            DBGrid1.Col = 4
            Cantidad = Val(DBGrid1.Text)
                    
            If Cantidad <> 0 Then
            
                WEstado = "N"
                SumaCant = 0
    
                WLote1 = xLote(Suma, 1)
                WLote2 = xLote(Suma, 3)
                Wlote3 = xLote(Suma, 5)
                WLote4 = xLote(Suma, 7)
                WLote5 = xLote(Suma, 9)
                WImpo = xLote(Suma, 2)
                WCanti1 = Str$(WImpo)
                WImpo = xLote(Suma, 4)
                WCanti2 = Str$(WImpo)
                WImpo = xLote(Suma, 6)
                WCanti3 = Str$(WImpo)
                WImpo = xLote(Suma, 8)
                WCanti4 = Str$(WImpo)
                WImpo = xLote(Suma, 10)
                WCanti5 = Str$(WImpo)
    
                If Val(WLote1) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti1)
                End If
                If Val(WLote2) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti2)
                End If
                If Val(Wlote3) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti3)
                End If
                If Val(WLote4) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti4)
                End If
                If Val(WLote5) <> 0 Then
                    SumaCant = SumaCant + Val(WCanti5)
                End If
    
                If SumaCant = Cantidad Then
                    WEstado = "S"
                        Else
                    WEstado = "N"
                    m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
                    A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                    Exit Sub
                End If
    
                If WEstado = "S" Then
    
                    Erase ControlLote
                    ControlLote(1, 1) = WLote1
                    ControlLote(1, 2) = WCanti1
                    ControlLote(2, 1) = WLote2
                    ControlLote(2, 2) = WCanti2
                    ControlLote(3, 1) = Wlote3
                    ControlLote(3, 2) = WCanti3
                    ControlLote(4, 1) = WLote4
                    ControlLote(4, 2) = WCanti4
                    ControlLote(5, 1) = WLote5
                    ControlLote(5, 2) = WCanti5
    
                    For Ciclo1 = 1 To 5
                        If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                            For Ciclo2 = 1 To 5
                                If Ciclo1 <> Ciclo2 Then
                                    If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                                        m$ = "A asignado una misma partida 2 veces"
                                        A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        WEstado = "N"
                                        Exit Sub
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
                    ControlLote(1, 1) = WLote1
                    ControlLote(1, 2) = WCanti1
                    ControlLote(2, 1) = WLote2
                    ControlLote(2, 2) = WCanti2
                    ControlLote(3, 1) = Wlote3
                    ControlLote(3, 2) = WCanti3
                    ControlLote(4, 1) = WLote4
                    ControlLote(4, 2) = WCanti4
                    ControlLote(5, 1) = WLote5
                    ControlLote(5, 2) = WCanti5
    
                    For Ciclo1 = 1 To 5
    
                        WLote = ControlLote(Ciclo1, 1)
                        WCanti = Val(ControlLote(Ciclo1, 2))
            
                        If Val(WLote) <> 0 Or Val(WCanti) <> 0 Then
            
                        If Left$(Articulo, 2) = "DY" Then
                            WTipoPro = "M"
                                Else
                            WTipoPro = "T"
                        End If
            
                        Select Case WTipoPro
                            Case "M"
                                WArti = Left$(Articulo, 3) + Right$(Articulo, 7)
                                WEntra = "N"
                                XParam = "'" + WLote + "','" _
                                             + WArti + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WSal = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Call Redondeo(WSal)
                                    WEntra = "S"
                                    If WSal < WCanti Then
                                        m$ = "La cantidad informada supera al saldo disponible"
                                        A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                        WEstado = "N"
                                        Exit Sub
                                    End If
                                    rstLaudo.Close
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArti + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        rstMovguia.Close
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                            Case Else
                                WEntra = "N"
                                WControla = 0
                                spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
                                    WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                                    rstTerminado.Close
                                End If
                
                                If WControla = 0 Then
                                    XParam = "'" + WLote + "','" _
                                            + Articulo + "'"
                                    spHoja = "ListaHojaProducto " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstHoja.RecordCount > 0 Then
                                        WSal = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                                        Call Redondeo(WSal)
                                        WEntra = "S"
                                        If WSal < WCanti Then
                                            m$ = "La cantidad informada supera al saldo disponible"
                                            A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                            WEstado = "N"
                                            Exit Sub
                                        End If
                                        rstHoja.Close
                                    End If
                
                                    If WEntra = "N" Then
                                        XParam = "'" + Articulo + "','" _
                                                    + WLote + "'"
                                        spMovguia = "ListaMovguiaLote1 " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstMovguia.RecordCount > 0 Then
                                            WSal = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                            Call Redondeo(WSal)
                                            WEntra = "S"
                                            If WSal < WCanti Then
                                                m$ = "La cantidad informada supera al saldo disponible"
                                                A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                                WEstado = "N"
                                                Exit Sub
                                            End If
                                            rstMovguia.Close
                                        End If
                                    End If
                
                                        Else
                                    WEntra = "S"
                                End If
                                If WEntra = "N" Then
                                    m$ = "Partida Inexistente"
                                    A = MsgBox(m$, 0, "PROBLEMAS EN LA ASIGNACION DE PARTIDAS")
                                    WEstado = "N"
                                    Exit Sub
                                End If
                
                        End Select
            
                        End If
            
                    Next Ciclo1

                End If
                
            End If
                                        
        Next iRow
            
    Next A
    
End Sub




