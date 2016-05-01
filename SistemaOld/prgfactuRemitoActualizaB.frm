VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactuRemitoActualizaB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision de Factura de Remitos ya emitidos"
   ClientHeight    =   8340
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
   Begin VB.CommandButton ConsultaPedido 
      Caption         =   "Consulta Remitos"
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
      Left            =   9120
      TabIndex        =   60
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Consignacion 
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
      Left            =   3000
      MaxLength       =   8
      TabIndex        =   58
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton salvacae 
      Caption         =   "salva cae"
      Height          =   375
      Left            =   7080
      TabIndex        =   57
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   56
      Text            =   " "
      Top             =   6120
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
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   54
      Text            =   " "
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11040
      TabIndex        =   52
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command1"
      Height          =   495
      Left            =   11280
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame PantaMotivo 
      Height          =   1815
      Left            =   480
      TabIndex        =   41
      Top             =   2400
      Visible         =   0   'False
      Width           =   10335
      Begin VB.ComboBox ConceptoAtraso 
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
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox DescriMotivo 
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
         MaxLength       =   50
         TabIndex        =   42
         Top             =   720
         Width           =   9855
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "MOTIVO DE RETRASO DE CUMPLIMIENTO DEL PEDIDO"
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
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   9735
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
      TabIndex        =   38
      Top             =   1200
      Width           =   2295
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
      TabIndex        =   33
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
      TabIndex        =   31
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   8760
      TabIndex        =   22
      Top             =   5640
      Width           =   2535
      Begin VB.Label Label26 
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
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label ImpoIbCiudad 
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
         TabIndex        =   46
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label ImpoIbTucu 
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
         TabIndex        =   45
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "IB Bs Aa"
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   30
         Top             =   2160
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
         TabIndex        =   29
         Top             =   1920
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
         TabIndex        =   28
         Top             =   1680
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   2160
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
         TabIndex        =   25
         Top             =   1920
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
         TabIndex        =   24
         Top             =   1680
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
         TabIndex        =   23
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
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   21
      Text            =   " "
      Top             =   120
      Width           =   975
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
      TabIndex        =   19
      Top             =   600
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
      Left            =   3360
      TabIndex        =   18
      Top             =   5640
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
      TabIndex        =   17
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
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   15
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   10
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6600
      TabIndex        =   8
      Top             =   120
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
      Left            =   960
      MaxLength       =   8
      TabIndex        =   6
      Text            =   " "
      Top             =   120
      Width           =   975
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
      TabIndex        =   4
      Top             =   0
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
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10560
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
      ItemData        =   "prgfactuRemitoActualizaB.frx":0000
      Left            =   4200
      List            =   "prgfactuRemitoActualizaB.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "prgfactuRemitoActualizaB.frx":0015
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
   Begin VB.Label Label21 
      Caption         =   "Consig."
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
      Left            =   2040
      TabIndex        =   59
      Top             =   120
      Width           =   855
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
      Left            =   6120
      TabIndex        =   55
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label14 
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
      Left            =   6120
      TabIndex        =   53
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   5640
      TabIndex        =   50
      Top             =   1200
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
      TabIndex        =   32
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
      Left            =   4080
      TabIndex        =   20
      Top             =   120
      Width           =   735
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
      TabIndex        =   16
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
      TabIndex        =   14
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   9
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
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "PrgFactuRemitoActualizaB"
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
Private WAdicional As Double
Private ZAdicional As String
Private WProvincia As String
Private WRubro As Integer
Private WVendedor As Integer
Private Precio As String
Private Dada As String
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
Private parcial As String

Private Auxiliar(100, 30) As String
Private RestaPedido(100, 3) As String
Private ClavePedido(100)

Private BajaLote(12, 2) As String
Dim CargaEmpresa(12, 2) As String

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
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstEstadisticaLote As Recordset
Dim spEstadisticaLote As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String
Dim rstCertificado As Recordset
Dim spCertificado As String

Dim XParam As String
Dim ZZImpreNumero As String

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

Dim ZZCampo1 As String
Dim ZZCampo2 As String

Dim WEstado As String
Dim XTerminado As String
Dim XCantidad  As Double
Dim WRow As Integer
Dim Compara As Double
Dim ZZIntervencion As String
Dim ZLugarFicha As Integer

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

Private WTipoPedido As String
Private WPorceIb As Double

Dim ZZFecha As String
Dim ZZDias As Integer
Dim ZZVto As String
Dim ZDolarEspecial As Integer

Dim ZLote1 As String
Dim ZCantidad1 As String
Dim ZLote2 As String
Dim ZCantidad2 As String
Dim ZLote3 As String
Dim ZCantidad3 As String
Dim ZLote4 As String
Dim ZCantidad4 As String
Dim ZLote5 As String
Dim ZCantidad5 As String
Dim ZLote6 As String
Dim ZCantidad6 As String
Dim ZLote7 As String
Dim ZCantidad7 As String
Dim ZLote8 As String
Dim ZCantidad8 As String
Dim ZLote9 As String
Dim ZCantidad9 As String
Dim ZLote10 As String
Dim ZCantidad10 As String
Dim ZLote11 As String
Dim ZCantidad11 As String
Dim ZLote12 As String
Dim ZCantidad12 As String

Dim ZEnv1 As String
Dim ZCantiEnv1 As String
Dim ZEnv2 As String
Dim ZCantiEnv2 As String
Dim ZEnv3 As String
Dim ZCantiEnv3 As String
Dim ZEnv4 As String
Dim ZCantiEnv4 As String
Dim ZEnv5 As String
Dim ZCantiEnv5 As String
Dim ZEnv6 As String
Dim ZCantiEnv6 As String
Dim ZEnv7 As String
Dim ZCantiEnv7 As String
Dim ZEnv8 As String
Dim ZCantiEnv8 As String
Dim ZEnv9 As String
Dim ZCantiEnv9 As String
Dim ZEnv10 As String
Dim ZCantiEnv10 As String
Dim ZEnv11 As String
Dim ZCantiEnv11 As String
Dim ZEnv12 As String
Dim ZCantiEnv12 As String

Dim WSal As Double
Dim WVector(10000, 4) As String
Dim ZClave  As String
Dim ZTipo As String
Dim ZNumero As String
Dim ZRenglon As String
Dim Renglon As Integer
Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim ZZValor1 As Double
Dim ZZValor2 As Double
Dim ZZVector(100, 5) As String
 
Dim DiaFeriado(100) As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim VectorCosto(100, 3) As String

Dim ZZLote As String

Dim ZMes As String
Dim ZAno As String
Dim ZClave1 As String
Dim ZClave2 As String
Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 4) As String
Dim ZDescri(10) As String
Dim ZDescriII(10) As String
Dim ZImpreFicha(100) As String

Dim ZZZProducto As String
Dim ZZZCosto As Double

Dim ZVersionPedido As Integer
Dim ZVersionAtraso As Integer
Dim ZSedronar As Integer
Dim ZNroSedronar As String

Dim ZZPasaImpre As Integer
Dim FF As Integer
Dim ZZGrabaFactura As String
Dim ZZImpreBarraI As String
Dim ZZImpreBarraII As String


Private Sub Calcula_Paridad()

    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZDolarEspecial = Trim(IIf(IsNull(rstCliente!DolarEspecial), "0", rstCliente!DolarEspecial))
        WPago2 = rstCliente!Pago2
        rstCliente.Close
    End If

    If ZDolarEspecial = 1 Then
        spPago = "ConsultaPago " + "'" + Str$(WPago2) + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            ZZDias = rstPago!Dias
        End If
        
        ZZFecha = Fecha.Text
        Call Calcula_vencimiento(ZZFecha, ZZDias, ZZVto)
        
        WMes = Val(Mid$(ZZVto, 4, 2))
        WAno = Val(Right$(ZZVto, 4))
        
        WMesII = Val(Mid$(ZZFecha, 4, 2))
        WAnoII = Val(Right$(ZZFecha, 4))
        
        
        For Ciclo = 1 To 4
            If WMesII = WMes And WAnoII = WAno Then
                Exit For
            End If
            WMesII = WMesII + 1
            If WMesII > 12 Then
                WAnoII = WAnoII + 1
                WMesII = 1
            End If
        Next Ciclo
        
        spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
        
            ZCambioIII = IIf(IsNull(rstCambios!CambioIII), "0", rstCambios!CambioIII)
            ZCambioIV = IIf(IsNull(rstCambios!CambioIV), "0", rstCambios!CambioIV)
            ZCambioV = IIf(IsNull(rstCambios!CambioV), "0", rstCambios!CambioV)
            ZCambioVI = IIf(IsNull(rstCambios!CambioVI), "0", rstCambios!CambioVI)
            
            Select Case Ciclo
                Case 1
                    Paridad.Text = Pusing("#,###.###", Str$(ZCambioIII))
                Case 2
                    Paridad.Text = Pusing("#,###.###", Str$(ZCambioIV))
                Case 3
                    Paridad.Text = Pusing("#,###.###", Str$(ZCambioV))
                Case Else
                    Paridad.Text = Pusing("#,###.###", Str$(ZCambioVI))
            End Select
            
            rstCambios.Close
                    Else
            Paridad.Text = ""
        End If
    
    End If
End Sub



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


Private Sub Calcula_Click()

    Call Calcula_Paridad
    
    WNeto = 0
    
    For a = 0 To 3
        
        Suma = a * 10
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
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 4
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    WImpoInteres = 0
    
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
    
    If Moneda.ListIndex = 0 Then
        Compara = WNeto * Val(Paridad.Text)
    End If
    Call Redondeo(Compara)
    If Compara < 100 Then
        WImpoIb = 0
    End If
    If Compara < 500 Then
        WImpoIbCiudad = 0
    End If
    
    Select Case Val(WCodIva)
        Case 2
            WIva1 = WNeto * 0.21
            WIva2 = WNeto * 0.105
            Call Redondeo(WIva1)
            Call Redondeo(WIva2)
        Rem Case 3, 4, 5
        Rem     WIva1 = 0
        Rem     WIva2 = 0
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

    With rstEmpresa
        .Close
    End With
    
    RetVal = Shell("cmd.exe /c Taskkill /f /IM AcroRd32.exe", 6)
    
    PrgFactuRemitoActualizaB.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    Rem XParam = "'" + "4336" + "'"
    Rem  , spEstadistica = "BajaEstadisticaNumero " + XParam
    Rem , Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem XParam = "'" + "4335" + "'"
    Rem spEstadistica = "BajaEstadisticaNumero " + XParam
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem Sql1 = "DELETE Estadistica"
    Rem Sql2 = " Where OrdFecha < " + "'" + "19990000" + "'"
    Rem spEstadistica = Sql1 + Sql2
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    WDesde = "19980101"
    WHasta = "19981231"
    Renglon = 0
    Erase WVector
    
    Sql1 = "Select *"
    Sql2 = " FROM CtaCte"
    Sql3 = " Where CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    Sql4 = " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    Sql5 = " and (CtaCte.Tipo = " + "'" + "01" + "'"
    Sql6 = " or CtaCte.Tipo = " + "'" + "02" + "'"
    Sql7 = " or CtaCte.Tipo = " + "'" + "03" + "'"
    Sql8 = " or CtaCte.Tipo = " + "'" + "04" + "'"
    Sql9 = " or CtaCte.Tipo = " + "'" + "05" + "')"
    spCtacte = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
        With rstCtacte
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    WVector(Renglon, 1) = rstCtacte!Clave
                    WVector(Renglon, 2) = rstCtacte!Tipo
                    WVector(Renglon, 3) = rstCtacte!Numero
                    WVector(Renglon, 4) = rstCtacte!Renglon
                    aa = rstCtacte!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCtacte.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        ZClaveII = WVector(Ciclo, 1)
        ZClave = WVector(Ciclo, 1)
        ZTipo = WVector(Ciclo, 2)
        ZNumero = Str$(Val(WVector(Ciclo, 3)) + 500000)
        ZRenglon = WVector(Ciclo, 4)
        
        Call Ceros(ZNumero, 8)
        ZClave = ZTipo + ZNumero + ZRenglon
        
        Sql1 = "UPDATE CtaCte SET "
        Sql2 = " Clave = " + "'" + ZClave + "',"
        Sql3 = " Numero = " + "'" + ZNumero + "'"
        Sql4 = " Where Clave = " + "'" + ZClaveII + "'"
                     
        spCtacte = Sql1 + Sql2 + Sql3 + Sql4
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
End Sub



Private Sub Command11_Click()
    Open "lpt1:" For Output As #1
    Rem Print #1, Chr$(27) & "&11h"
    Print #1, Chr$(27) & "&l4H"
    Print #1, "bandeja 1"
    Print #1, Chr$(27) & "&l1H"
    Print #1, "bandeja 2"
    Print #1, Chr$(12)
    Close #1


End Sub

Private Sub Consignacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Auxi = Consignacion.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "10" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
        
            Pedido.Text = rstCtacte!Pedido
            Rem Fecha.Text = rstCtacte!Fecha
            Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            Cliente.Text = rstCtacte!Cliente
            Remito.Text = rstCtacte!Remito
            
            rstCtacte.Close
            
            spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                rstPedido.Close
            End If
                
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
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
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                Rem WDirentrega = rstCliente!DirEntrega
                WDirentrega = ""
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                rstCliente.Close
            End If
            
            Call Pedido_KeyPress(13)
            Call Fecha_Keypress(13)
            
            Call Calcula_Paridad
            Call Proceso2_Click
            Call Calcula_Click
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 4
            DBGrid1.Row = 0
            
            Orden.SetFocus
            
        End If
    End If


End Sub

Private Sub Descri4_Click()
End Sub

Private Sub ConsultaPedido_Click()
    ZZProcesoFactura = 1
    PrgSeleccionaRemito.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    If Val(WEmpresa) = 1 Then
        OPEN_FILE_Ctacte8
        OPEN_FILE_Numero8
        OPEN_FILE_Esta8
    End If
    
    If ZZProcesoFactura = 99 And Val(Consignacion.Text) <> 0 Then
        Call Consignacion_KeyPress(13)
        Call Fecha_Keypress(13)
        Call Calcula_Click
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 5
        DBGrid1.Row = 0
        Consignacion.SetFocus
    End If
    
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    If Val(Paridad.Text) = 0 Then
        Exit Sub
    End If
    
    If Trim(Cae.Text) = "" Then
        ZZGrabaFactura = ""
        Call Calcula_Cae
        If ZZGrabaFactura <> "S" Then
            Exit Sub
        End If
    End If
    
    Call Calcula_Click
    
    ZSql = ""
    ZSql = ZSql + "DELETE CtaCte"
    ZSql = ZSql + " Where Ctacte.Tipo = " + "'" + "10" + "'"
    ZSql = ZSql + " and Ctacte.Numero = " + "'" + Consignacion.Text + "'"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    WTipo = "01"
    WNumero = Numero.Text
    WRenglon = "01"
    WCliente = Cliente.Text
    WFecha = Fecha.Text
    WEstado = "0"
    Call Convierte_datos(Str$(Total), Auxi)
    
    If Moneda.ListIndex = 0 Then
        XTotalUs = Str$(WTotal)
        XTotal = Str$(WTotal * Val(Paridad.Text))
        XSaldoUs = Str$(WTotal)
        XSaldo = Str$(WTotal * Val(Paridad.Text))
        XNet = Str$(WNeto * Val(Paridad.Text))
        XIva1 = Str$(WIva1 * Val(Paridad.Text))
        XIva2 = Str$(WIva2 * Val(Paridad.Text))
        XImpoIb = Str$(WImpoIb * Val(Paridad.Text))
        XImpoIbTucu = Str$(WImpoIbTucu * Val(Paridad.Text))
        XImpoIbCiudad = Str$(WImpoIbCiudad * Val(Paridad.Text))
            Else
        XTotalUs = Str$(WTotal / Val(Paridad.Text))
        XTotal = Str$(WTotal)
        XSaldoUs = Str$(WTotal / Val(Paridad.Text))
        XSaldo = Str$(WTotal)
        XNet = Str$(WNeto)
        XIva1 = Str$(WIva1)
        XIva2 = Str$(WIva2)
        XImpoIb = Str$(WImpoIb)
        XImpoIbTucu = Str$(WImpoIbTucu)
        XImpoIbCiudad = Str$(WImpoIbCiudad)
    End If
        
    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    WOrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
    WOrdVencimiento1 = Right$(WVencimiento1, 4) + Mid$(WVencimiento1, 4, 2) + Left$(WVencimiento1, 2)
    WImpre = "FC"
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
    
    ZZImpreNumero = Str$(Val(WNumero) - 200000)
    Call Ceros(ZZImpreNumero, 8)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql + " Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql + " fechaCae = " + "'" + vtocae.Text + "',"
    ZSql = ZSql + " Moneda = " + "'" + Str$(Moneda.ListIndex) + "',"
    ZSql = ZSql + " ImpoIbTucu = " + "'" + XImpoIbTucu + "',"
    ZSql = ZSql + " ImpoIbCiudad = " + "'" + XImpoIbCiudad + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                 
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    WAdicional = 0
    WNumero8 = ""
    WImporte8 = 0
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
        rstCliente.Close
    End If
    
    ZCliente = Cliente.Text
    
    If WAdicional > 0 Then
        If Val(WEmpresa) = 8 Then
            OPEN_FILE_Ctacte8
            OPEN_FILE_Numero8
            OPEN_FILE_Esta8
            If Cliente.Text = "S00016" Then
                ZCliente = "A00013"
            End If
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
        
        If Moneda.ListIndex = 0 Then
            WImporte8 = (WNeto * WAdicional) * Val(Paridad.Text)
                Else
            WImporte8 = (WNeto * WAdicional)
        End If
        
        With rstCtacte8
            .Index = "Clave"
            .AddNew
            !Tipo = "01"
            !Numero = WNumero8
            !Renglon = "00"
            !Cliente = ZCliente
            !Fecha = Fecha.Text
            !Estado = "0"
            !Vencimiento = "  /  /    "
            !Vencimiento1 = "  /  /    "
            Call Convierte_datos(Str$(Total), Auxi)
            If Moneda.ListIndex = 0 Then
                !Total = (WNeto * WAdicional) * Val(Paridad.Text)
                !Totalus = (WNeto * WAdicional)
                !Saldo = (WNeto * WAdicional) * Val(Paridad.Text)
                !Saldous = (WNeto * WAdicional)
                    Else
                !Total = (WNeto * WAdicional)
                !Totalus = (WNeto * WAdicional) / Val(Paridad.Text)
                !Saldo = (WNeto * WAdicional)
                !Saldous = (WNeto * WAdicional) / Val(Paridad.Text)
            End If
            !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            !OrdVencimiento = "00000000"
            !OrdVencimiento1 = "00000000"
            !Impre = "FC"
            If Moneda.ListIndex = 0 Then
                !Neto = (WNeto * WAdicional) * Val(Paridad.Text)
                    Else
                !Neto = (WNeto * WAdicional)
            End If
            !Iva1 = 0
            !Iva2 = 0
            !Pedido = 0
            !Remito = 0
            !Orden = ""
            !Paridad = Val(Paridad.Text)
            !Provincia = WProv
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
            !Clave = "01" + Auxi + "00"
            !WDate = Date$
            !TipoDescarga = 1
            .Update
        End With
        
    End If
    
    ZAdicional = Str$(WAdicional)
    ZAdicional = Pusing("######", ZAdicional)
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = "01" + Auxi + "01"
    
    
    Sql1 = "UPDATE Ctacte SET "
    Sql2 = "Adicional = " + "'" + ZAdicional + "',"
    Sql3 = "Numero8 = " + "'" + WNumero8 + "'"
    Sql4 = " Where Clave = " + "'" + WClave + "'"
                 
    spCtacte = Sql1 + Sql2 + Sql3 + Sql4
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    If WAdicional > 0 Then
    
        Auxi = WNumero8
        Call Ceros(Auxi, 8)
        WClave = "01" + Auxi + "00"
        With rstCtacte8
            .Index = "Clave"
            .Seek "=", WClave
            If .NoMatch = True Then
                m$ = "No se ha podido generar correctamente la factura complementaria"
                a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    Else
                If WImporte8 <> !Total Then
                    m$ = "No se ha podido generar correctamente la factura complementaria"
                    a% = MsgBox(m$, 0, "MODULO DE FACTURACION")
                End If
            End If
        End With
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZVector
    ZZLugarII = 0
    
    
    Suma = 0
    Renglon = 0
    Renglon1 = 0
    WRenglon = 0
    DBGrid1.Refresh
    
    For a = 0 To 3
    
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        
        For iRow = 0 To 9
        
            Suma = Suma + 1
            WRenglon = WRenglon + 1
            
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
            WTipoProDy = Left$(Articulo, 2)
                
            DBGrid1.Col = 1
            ZZDescriArticulo = DBGrid1.Text
                
            DBGrid1.Col = 3
            ZZPrecio = Val(DBGrid1.Text)
                
            DBGrid1.Col = 4
            ZZCantidad = Val(DBGrid1.Text)
            
            DBGrid1.Col = 5
            ZZRestaCantidad = Val(DBGrid1.Text)
                
            If ZZCantidad <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Val(Consignacion.Text) + 900000)
                Call Ceros(Auxi1, 8)
                WClaveAnt = "01" + Auxi1 + Auxi
                
                Auxi1 = Numero.Text
                Call Ceros(Auxi1, 8)
                WClave = "01" + Auxi1 + Auxi
                
                WTipo = "01"
                WNumero = Numero.Text
                If Moneda.ListIndex = 0 Then
                    XPrecioUs = Str$(ZZPrecio)
                    XPrecio = Str$(ZZPrecio * Val(Paridad.Text))
                    XImporteUs = Str$(ZZPrecio * ZZCantidad)
                    XImporte = Str$(ZZPrecio * Val(Paridad.Text) * ZZCantidad)
                        Else
                    XPrecioUs = Str$(ZZPrecio / Val(Paridad.Text))
                    XPrecio = Str$(ZZPrecio)
                    XImporteUs = Str$((ZZPrecio * ZZCantidad) / Val(Paridad.Text))
                    XImporte = Str$(ZZPrecio * ZZCantidad)
                End If
                
                WParidad = Paridad.Text

                ZSql = ""
                ZSql = ZSql + "UPDATE Estadistica SET "
                ZSql = ZSql + " Precio = " + "'" + XPrecio + "',"
                ZSql = ZSql + " PrecioUs = " + "'" + XPrecioUs + "',"
                ZSql = ZSql + " Importe = " + "'" + XImporte + "',"
                ZSql = ZSql + " ImporteUs = " + "'" + XImporteUs + "',"
                ZSql = ZSql + " Paridad = " + "'" + WParidad + "',"
                ZSql = ZSql + " Numero = " + "'" + Numero.Text + "',"
                ZSql = ZSql + " Clave = " + "'" + WClave + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClaveAnt + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                ZZLugarII = ZZLugarII + 1
            
                ZZVector(ZZLugarII, 1) = Str$(ZZCantidad)
                ZZVector(ZZLugarII, 2) = ZZDescriArticulo
                ZZVector(ZZLugarII, 3) = Str$(ZZPrecio)
                ZZVector(ZZLugarII, 4) = Str$(ZZPrecio * ZZCantidad)
                        
                WTipoProDy = Left$(Articulo, 2)
                If WTipoProDy <> "PT" Then
                    XTipoproDy = "M"
                    XArticuloDy = Left$(Articulo, 3) + Right$(Articulo, 7)
                        Else
                    XTipoproDy = "T"
                    XArticuloDy = "  -   -   "
                End If
                        
                If XTipoproDy = "M" Then
                
                    ClavePrecioMp = Cliente.Text + XArticuloDy
                
                    spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePrecioMp + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
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
                    
                        If rstPreciosMp!Cantidad2 <> O Then
                            WFecha1 = rstPreciosMp!Fecha2
                            WFactura1 = rstPreciosMp!Factura2
                            WPrecio1 = Str$(rstPreciosMp!Precio2)
                            WCantidad1 = Str$(rstPreciosMp!Cantidad2)
                        End If
                                    
                        If rstPreciosMp!Cantidad3 <> O Then
                            WFecha2 = rstPreciosMp!Fecha3
                            WFactura2 = rstPreciosMp!Factura3
                            WPrecio2 = Str$(rstPreciosMp!Precio3)
                            WCantidad2 = Str$(rstPreciosMp!Cantidad3)
                        End If
                                    
                        If rstPreciosMp!Cantidad4 <> O Then
                            WFecha3 = rstPreciosMp!Fecha4
                            WFactura3 = rstPreciosMp!Factura4
                            WPrecio3 = Str$(rstPreciosMp!Precio4)
                            WCantidad3 = Str$(rstPreciosMp!Cantidad4)
                        End If
                                    
                        If rstPreciosMp!Cantidad5 <> O Then
                            WFecha4 = rstPreciosMp!Fecha5
                            WFactura4 = rstPreciosMp!Factura5
                            WPrecio4 = Str$(rstPreciosMp!Precio5)
                            WCantidad4 = Str$(rstPreciosMp!Cantidad5)
                        End If
                                    
                        WFecha5 = Fecha.Text
                        WFactura5 = Numero.Text
                        If Moneda.ListIndex = 0 Then
                            WPrecio5 = Str$(ZZPrecio)
                                Else
                            WPrecio5 = Str$(ZZPrecio / Val(Paridad.Text))
                        End If
                        WCantidad5 = Str$(ZZCantidad)
                                    
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
                            WFecha1 = rstPrecios!Fecha2
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
                        WPrecio5 = Str$(ZZPrecio)
                        WCantidad5 = Str$(ZZCantidad)
                                    
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
                
                
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    If Cliente.Text <> "P00005" Then

                        Select Case WTipoPedido
                            Case "FA", "PT", "BI", "TA"
                        
                                XEmpresa = WEmpresa
                                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                    Select Case WTipoPedido
                                        Case "PG", "CO"
                                            WEmpresa = "0001"
                                            txtOdbc = "Empresa01"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "FA"
                                            WEmpresa = "0011"
                                            txtOdbc = "Empresa11"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case "TA"
                                            WEmpresa = "0003"
                                            txtOdbc = "Empresa03"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        Case Else
                                            WEmpresa = "0007"
                                            txtOdbc = "Empresa07"
                                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    End Select
                                End If
                        
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Estadistica SET "
                                ZSql = ZSql + " Precio = " + "'" + XPrecio + "',"
                                ZSql = ZSql + " PrecioUs = " + "'" + XPrecioUs + "',"
                                ZSql = ZSql + " Importe = " + "'" + XImporte + "',"
                                ZSql = ZSql + " ImporteUs = " + "'" + XImporteUs + "',"
                                ZSql = ZSql + " Paridad = " + "'" + WParidad + "',"
                                ZSql = ZSql + " Clave = " + "'" + WClave + "'"
                                ZSql = ZSql + " Where Clave = " + "'" + WClaveAnt + "'"
                                spEstadistica = ZSql
                                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                
                                Call Conecta_Empresa
                            
                            Case Else
                            
                        End Select
                        
                            Else
                        
                        If Left$(WArticulo, 4) <> "PT-5" Then
                        
                            Select Case WTipoPedido
                                Case "FA", "PT", "BI", "TA"
                        
                                    XEmpresa = WEmpresa
                                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                        Select Case WTipoPedido
                                            Case "PG", "CO"
                                                WEmpresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case "FA"
                                                WEmpresa = "0011"
                                                txtOdbc = "Empresa11"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case "TA"
                                                WEmpresa = "0003"
                                                txtOdbc = "Empresa03"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                                WEmpresa = "0007"
                                                txtOdbc = "Empresa07"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        End Select
                                    End If
                            
                                    ZSql = ""
                                    ZSql = ZSql + "UPDATE Estadistica SET "
                                    ZSql = ZSql + " Precio = " + "'" + XPrecio + "',"
                                    ZSql = ZSql + " PrecioUs = " + "'" + XPrecioUs + "',"
                                    ZSql = ZSql + " Importe = " + "'" + XImporte + "',"
                                    ZSql = ZSql + " ImporteUs = " + "'" + XImporteUs + "',"
                                    ZSql = ZSql + " Paridad = " + "'" + WParidad + "',"
                                    ZSql = ZSql + " Clave = " + "'" + WClave + "'"
                                    ZSql = ZSql + " Where Clave = " + "'" + WClaveAnt + "'"
                                    spEstadistica = ZSql
                                    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                                
                                    Call Conecta_Empresa
                            
                                Case Else
                                
                            End Select
                            
                        End If
                        
                    End If
                End If
                
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
                        !Cantidad = ZZCantidad
                        If Moneda.ListIndex = 0 Then
                            !Precio = ZZPrecio * Val(Paridad.Text) * WAdicional
                            !PrecioUs = ZZPrecio * WAdicional
                            !Importe = ZZPrecio * ZZCantidad * Val(Paridad.Text) * WAdicional
                            !ImporteUs = ZZPrecio * ZZCantidad * WAdicional
                                Else
                            !Precio = ZZPrecio * WAdicional
                            !PrecioUs = ZZPrecio * WAdicional / Val(Paridad.Text)
                            !Importe = ZZPrecio * ZZCantidad * WAdicional
                            !ImporteUs = ZZPrecio * ZZCantidad * WAdicional / Val(Paridad.Text)
                        End If
                        !Cliente = ZCliente
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
                
            End If
                                    
        Next iRow
        
    Next a

    spNumero = "ConsultaNumero " + "'" + "61" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        WCodigo = "61"
        WNumero = Numero.Text
        rstNumero.Close
        XParam = "'" + WCodigo + "','" _
                     + WNumero + "'"
        spNumero = "ModificaNumero " + XParam
        Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Call ImpresionFe
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    PrgFactuRemitoActualizaB.Show
    Numero.SetFocus
        
    Exit Sub

WError:
    MsgBox Err.Description
    Resume Next
        
End Sub

Private Sub Verifica_Fecha_Entrega()

    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZTipoPedido = rstPedido!Tipoped
        ZFecha = rstPedido!Fecha
        ZFechaEntrega = rstPedido!FecEntrega
        ZOrdFechaEntrega = rstPedido!OrdFecEntrega
        ZFechaActualizacion = IIf(IsNull(rstPedido!FechaActualizacion), "", rstPedido!FechaActualizacion)
        ZOrdFechaActualizacion = IIf(IsNull(rstPedido!OrdFechaActualizacion), "", rstPedido!OrdFechaActualizacion)
        rstPedido.Close
    End If
    
    If ZTipoPedido = 4 Then
        If ZFechaActualizacion <> "" Then
            ZFechaFactu = ZFechaActualizacion
            ZFechaFactuOrd = ZOrdFechaActualizacion
                Else
            ZFechaFactu = Fecha.Text
            ZFechaFactuOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        End If
                Else
        ZFechaFactu = Fecha.Text
        ZFechaFactuOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    End If
        
    WDias = 0
    WSuma2 = "0"
    WFechaHastaOrd = ZFechaFactuOrd
    WFechaDesdeOrd = ZOrdFechaEntrega
    WFechaHasta = ZFechaFactu
    WFechaDesde = ZFechaEntrega
            
    If WFechaHastaOrd > WFechaDesdeOrd Then
            
        WSuma2 = "1"
            
        Do
        
            Feriado = "N"
            For Cicla = 1 To TotalFeriado
                If DiaFeriado(Cicla) = WFechaDesde Then
                    Feriado = "S"
                    Exit For
                End If
            Next Cicla
                    
            Rem 1 - DOMINGO
            Rem 2 - LUNES
            Rem 3 - MARTES
            Rem 4 - MIERCOLES
            Rem 5 - JUEVES
            Rem 6 - VIERNES
            Rem 7 - SABADO
            XFec1 = WFechaDesde
            strDia = Format$(XFec1, "dddd")
            BDia = Format(XFec1, "w")
            If BDia = 1 Or BDia = 7 Then
                Feriado = "S"
            End If
            
            If Feriado = "N" Then
                WDias = WDias + 1
            End If
            SumaDia = 2
            Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
            WFechaDesde = XFec2
                        
            If WFechaDesde = WFechaHasta Then
                Exit Do
            End If
        
        Loop
        
    End If
    
    Fecha.SetFocus
    
    If WDias > 0 Then
    
        ZVersionAtraso = 0
        ZVersionPedido = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Atraso"
        ZSql = ZSql + " Where Atraso.Pedido = " + "'" + Pedido.Text + "'"
        ZSql = ZSql + " Order by Atraso.Numero"
        spAtraso = ZSql
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstAtraso.RecordCount > 0 Then
            With rstAtraso
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        ZOrigen = IIf(IsNull(!Origen), "0", !Origen)
                        If ZOrigen = 0 Then
                            ZVersionAtraso = IIf(IsNull(!VersionPedido), "0", !VersionPedido)
                        End If
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstAtraso.Close
        End If
        
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZVersionPedido = rstPedido!Version
            rstPedido.Close
        End If
            
        If ZVersionPedido <> ZVersionAtraso Then
            ConceptoAtraso.ListIndex = 0
            DescriMotivo.Text = ""
            PantaMotivo.Visible = True
            DescriMotivo.SetFocus
        End If
        
    End If

End Sub

Private Sub DescriMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(DescriMotivo.Text)) >= 10 And ConceptoAtraso.ListIndex > 0 Then
        
            ZZAtraso = "1"
        
            Sql1 = "Select Max(Numero) as [NumeroMayor]"
            Sql2 = " FROM Atraso"
            spAtraso = Sql1 + Sql2
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                ZZAtraso = Str$(rstAtraso!Numeromayor + 1)
                rstAtraso.Close
            End If
    
            ZFecha = Fecha.Text
            ZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZFechaEntrega = Fecha.Text
            ZFechaEntregaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZTerminado = "  -     -   "
            ZArticulo = "  -   -   "
            ZDesTerminado = ""
            ZDesArticulo = ""
            ZConcepto = Str$(ConceptoAtraso.ListIndex + 4)
            ZSolicitud = ""
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Atraso ("
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Problema ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "FechaEntrega ,"
            ZSql = ZSql + "OrdFechaEntrega ,"
            ZSql = ZSql + "DesCliente ,"
            ZSql = ZSql + "DesTerminado ,"
            ZSql = ZSql + "DesArticulo ,"
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Solicitud ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "VersionPedido)"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZAtraso + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZFechaOrd + "',"
            ZSql = ZSql + "'" + Pedido.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + DescriMotivo.Text + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZFechaEntrega + "',"
            ZSql = ZSql + "'" + ZFechaEntregaOrd + "',"
            ZSql = ZSql + "'" + Left$(DesCliente.Caption, 50) + "',"
            ZSql = ZSql + "'" + ZDesTerminado + "',"
            ZSql = ZSql + "'" + ZDesArticulo + "',"
            ZSql = ZSql + "'" + ZConcepto + "',"
            ZSql = ZSql + "'" + ZSolicitud + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + "" + "')"
           
            spAtraso = ZSql
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        
            PantaMotivo.Visible = False
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    Cae.Text = ""
    vtocae.Text = "00/00/0000"
    Moneda.ListIndex = 1
    Consignacion.Text = ""
    
    For a = 0 To 3
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
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    WAdicional = 0
    
    spNumero = "ConsultaNumero " + "'" + "61" + "'"
    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
    If rstNumero.RecordCount > 0 Then
        Numero.Text = rstNumero!Numero + 1
        rstNumero.Close
            Else
        Numero.Text = ""
    End If
    
    spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
        rstCambios.Close
                Else
        Paridad.Text = ""
    End If
    
    Rem Numero.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 6
                If DBGrid1.Row < 40 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    WRow = DBGrid1.Row
                    DBGrid1.Col = 4
                    KeyCode = 0
                End If
            
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
    
    Tipoventa.Clear
    
    Tipoventa.AddItem "Venta Normal"
    Tipoventa.AddItem "Mercaderia en Consignacion"
    
    Tipoventa.ListIndex = 0
    
    ConceptoAtraso.Clear
    
    ConceptoAtraso.AddItem ""
    ConceptoAtraso.AddItem "Error del Sistema"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Problemas Vehiculos"
    ConceptoAtraso.AddItem "Problemas Logistica"
    ConceptoAtraso.AddItem "Problemas Recepcion Cliente"
    ConceptoAtraso.AddItem "Varios"
    ConceptoAtraso.AddItem "Corte de Luz"
    ConceptoAtraso.AddItem "Pedido por el Cliente"
    ConceptoAtraso.AddItem "Falta de Pago"
    ConceptoAtraso.AddItem "Confirmacion Pedido Parcial"
    ConceptoAtraso.AddItem "Envase"
    
    ConceptoAtraso.ListIndex = 0
    
    
    Moneda.Clear
    
    Moneda.AddItem "U$S"
    Moneda.AddItem "Pesos"
    
    Moneda.ListIndex = 1

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

    Numero.Text = ""
    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Remito.Text = ""
    Orden.Text = ""
    Cae.Text = ""
    vtocae.Text = "00/00/0000"
    Consignacion.Text = ""
    
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    ImpoIb.Caption = ""
    ImpoIbTucu.Caption = ""
    ImpoIbCiudad.Caption = ""
    Total.Caption = ""
    Paridad.Text = ""
    Dto.Caption = ""
    Interes.Caption = ""
    WAdicional = 0
    
    Renglon = 0
    
    spNumero = "ConsultaNumero " + "'" + "61" + "'"
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
    
    spCambios = "ConsultaCambio " + "'" + Fecha.Text + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
        rstCambios.Close
                Else
        Paridad.Text = ""
    End If
    
    Numero.SetFocus
     
End Sub

Private Sub Proceso2_Click()

    WNeto = 0

    For a = 0 To 3
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
    Erase Auxiliar
    
    XConsig = Str$(Val(Consignacion.Text) + 900000)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where EStadistica.Tipo = " + "'" + "01" + "'"
    ZSql = ZSql + " and Estadistica.Numero = " + "'" + XConsig + "'"
    spEstadistica = ZSql
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
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    If Moneda.ListIndex = 0 Then
                        Dada = Str$(rstEstadistica!PrecioUs)
                            Else
                        Dada = Str$(rstEstadistica!Precio)
                    End If
                    Dada = ""
                    DBGrid1.Col = 3
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    Dada = Str$(rstEstadistica!Cantidad)
                    DBGrid1.Col = 4
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                    
                    DBGrid1.Col = 5
                    DBGrid1.Text = Pusing("###,###.##", Dada)
                
                    If !Cantidad <> 0 Then
                        WNeto = WNeto + (rstEstadistica!Cantidad * rstEstadistica!Precio)
                    End If
                    
                    Auxiliar(Renglon, 1) = Auxi1
                    Auxiliar(Renglon, 2) = Str$(!Cantidad)
                    Auxiliar(Renglon, 3) = Str$(!Precio)
                    
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
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
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
                    
                    Auxiliar(DA, 30) = rstArticulo!Descripcion
                    
                    rstArticulo.Close
                    
                    WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                    ClavePreciosMp = Cliente.Text + WArti
                
                    spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPreciosMp.RecordCount > 0 Then
                
                        DBGrid1.Col = 3
                        If Moneda.ListIndex = 0 Then
                            DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio))
                            Precio = rstPreciosMp!Precio
                                Else
                            DBGrid1.Text = Pusing("###,###.##", Str$(rstPreciosMp!Precio * Val(Paridad.Text)))
                            Precio = rstPreciosMp!Precio * Val(Paridad.Text)
                        End If
                
                        rstPreciosMp.Close
                    End If
                    
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
                    
                    DBGrid1.Col = 3
                    If Moneda.ListIndex = 0 Then
                        DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
                        Precio = rstPrecios!Precio
                            Else
                        DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio * Val(Paridad.Text)))
                        Precio = rstPrecios!Precio * Val(Paridad.Text)
                    End If
                    
                    Auxiliar(DA, 30) = rstPrecios!Descripcion
                    
                    rstPrecios.Close
                                
                End If
        End Select
        
    Next DA
    
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
    
    Rem Graba.Enabled = False
    Rem Borra.Enabled = False

End Sub



Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        ClaveCtacte = "01" + Auxi + "01"
    
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            rstCtacte.Close
            m$ = "Factura ya existente"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Exit Sub
            
                    Else
                    
            WNumero = Numero.Text
            Consignacion.SetFocus
                
        End If
    End If
End Sub


Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            Orden.Text = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
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
            
            If Val(WEmpresa) = 1 And Cliente.Text = "P00005" Then
                If Left$(rstPedido!Terminado, 4) = "PT-5" Or rstPedido!Terminado = "PT-03000-001" Then
                    WTipoPedido = "PG"
                End If
            End If
            
            If Left$(rstPedido!Terminado, 4) = "PT-4" Then
                WTipoPedido = "TA"
            End If
            
            rstPedido.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                WPago1 = rstCliente!Pago1
                WPago2 = rstCliente!Pago2
                WVendedor = rstCliente!vendedor
                WRubro = rstCliente!Rubro
                WCodIva = rstCliente!Iva
                WAdicional = IIf(IsNull(rstCliente!Adicional), "0", rstCliente!Adicional)
                WCodIb = rstCliente!Ib
                WCodIbTucu = IIf(IsNull(rstCliente!IbTucu), "0", rstCliente!IbTucu)
                WCodIbCiudad = IIf(IsNull(rstCliente!IbCiudad), "0", rstCliente!IbCiudad)
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WProv = rstCliente!Provincia
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                Rem WDirentrega = rstCliente!DirEntrega
                WDirentrega = ""
                ZDirEntrega(1) = rstCliente!DirEntrega
                ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
                WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                rstCliente.Close
            End If
            Call Calcula_Paridad
            Call Calcula_FechaVto
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
                Paridad.Text = Pusing("#,###.###", Str$(rstCambios!Cambio))
                        Else
                Paridad.Text = ""
                Rem m$ = "No exsite paridad cargada para esta fecha"
                Rem a% = MsgBox(m$, 0, "Emision de facturas")
                Rem Fecha.SetFocus
            End If
            If Val(Paridad.Text) <> 0 Then
                Call Calcula_FechaVto
                Vencimiento.Text = Wvencimiento
                
                Remito.SetFocus
                    Else
                m$ = "No exsite paridad cargada para esta fecha"
                a% = MsgBox(m$, 0, "Emision de facturas")
                Fecha.SetFocus
            End If
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de facturas")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub reImpre_Click()

    T$ = "Impresion"
    m$ = "Desea imprimir la factura electronica"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call ImpresionFe
    End If
    
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
End Sub

Private Sub salvacae_Click()

    Call Calcula_Cae
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WClave = "01" + Auxi + "01"
    
    ZZImpreNumero = Str$(Val(WNumero) - 200000)
    Call Ceros(ZZImpreNumero, 8)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " ImpreNumero = " + "'" + ZZImpreNumero + "',"
    ZSql = ZSql + " Cae = " + "'" + Cae.Text + "',"
    ZSql = ZSql + " fechaCae = " + "'" + vtocae.Text + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                 
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
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

Sub ImpresionFe()

    Call Calcula_Barra
    
    Auxi1 = Str$(Val(Numero.Text) - 200000)
    Call Ceros(Auxi1, 8)
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpreFactura"
    Rem ZSql = ZSql + " Where Numero = " + "'" + Auxi1 + "'"
    spImpreFactura = ZSql
    Set rstImpreFactura = db.OpenRecordset(spImpreFactura, dbOpenSnapshot, dbSQLPassThrough)

    ImporteIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
    ImporteIbTucu = Val(ImpoIbTucu.Caption) * Val(Paridad.Text)
    ImporteIbCiudad = Val(ImpoIbCiudad.Caption) * Val(Paridad.Text)
    ImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
    ImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
    Impotot = Val(Total.Caption) * Val(Paridad.Text)
    
    Impre = 0
    ImpreDespachoI = ""
    ImpreDespachoII = ""
        
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
        WDirentrega = ""
        ZDirEntrega(1) = rstCliente!DirEntrega
        ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        WDirentrega = ZDirEntrega(ZLugarDirEntrega)
        rstCliente.Close
    End If
        
    If Moneda.ListIndex = 0 Then
        ZZImpoIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
        ZZImpoIbTucu = Val(ImpoIbTucu.Caption) * Val(Paridad.Text)
        ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption) * Val(Paridad.Text)
        ZZImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
        ZZImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
        ZZImpoTotal = Val(Total.Caption) * Val(Paridad.Text)
            Else
        ZZImpoIb = Val(ImpoIb.Caption)
        ZZImpoIbTucu = Val(ImpoIbTucu.Caption)
        ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption)
        ZZImpoNeto = Val(Neto.Caption)
        ZZImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption))
        ZZImpoTotal = Val(Total.Caption)
    End If
        

    For aa = Impre To 16
                
        ZZCantidad = Val(ZZVector(aa, 1))
        ZZDescripcion = ZZVector(aa, 2)
        ZZPrecio = Val(ZZVector(aa, 3)) * 1.21
        ZZParcial = Val(ZZVector(aa, 4)) * 1.21
                    
        If Val(Numero.Text) > 200000 Then
            Auxi1 = Str$(Val(Numero.Text) - 200000)
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
        WWOrden = Trim(Orden.Text)
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
        WWImpreComprobante = "FACTURA"
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
        Rem by nan
      Rem  wwpago = ""
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
    
    If Val(WEmpresa) = 1 Then
        If Moneda.ListIndex = 0 Then
            Listado.ReportFileName = "ImpreFacturaLocalDolarNuevob.rpt"
                Else
            Listado.ReportFileName = "ImpreFacturaLocalPesosNuevob.rpt"
        End If
            Else
        If Moneda.ListIndex = 0 Then
            Listado.ReportFileName = "ImpreFacturaLocalDolarPellib.rpt"
                Else
            Listado.ReportFileName = "ImpreFacturaLocalPesosPellib.rpt"
        End If
    End If
                
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
    Select Case Val(WEmpresa)
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
    tipo_cbte = 6
    Select Case Val(WEmpresa)
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
    
    If Val(Numero.Text) - 200000 <> Cbte_Nro + 1 Then
        m$ = "El numero de comprobante no es igua al correlativo indicado por la afip " + Str$(LastCBTE + 1)
        a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
        Exit Sub
    End If
    
    WImpoIb = 0
    WImpoIbTucu = 0
    WImpoIbCiudad = 0
    
    If Moneda.ListIndex = 0 Then
        Rem MONEDA
        moneda_id = "DOL"
        Rem COTIZACION
        moneda_ctz = Val(Paridad.Text)
        ZZImpoIb = Val(ImpoIb.Caption)
        ZZImpoIbTucu = Val(ImpoIbTucu.Caption)
        ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption)
        ZZImpoNeto = Val(Neto.Caption)
        ZZImpoIva = Val(Iva1.Caption) + Val(Iva2.Caption)
        ZZImpoTotal = Val(Total.Caption)
        Rem ZZImpoIb = Val(ImpoIb.Caption) * Val(Paridad.Text)
        Rem ZZImpoIbTucu = Val(ImpoIbTucu.Caption) * Val(Paridad.Text)
        Rem ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption) * Val(Paridad.Text)
        Rem ZZImpoNeto = Val(Neto.Caption) * Val(Paridad.Text)
        Rem ZZImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption)) * Val(Paridad.Text)
        Rem ZZImpoTotal = Val(Total.Caption) * Val(Paridad.Text)
            Else
        Rem MONEDA
        moneda_id = "PES"
        Rem COTIZACION
        moneda_ctz = 1
        ZZImpoIb = Val(ImpoIb.Caption)
        ZZImpoIbTucu = Val(ImpoIbTucu.Caption)
        ZZImpoIbCiudad = Val(ImpoIbCiudad.Caption)
        ZZImpoNeto = Val(Neto.Caption)
        ZZImpoIva = (Val(Iva1.Caption) + Val(Iva2.Caption))
        ZZImpoTotal = Val(Total.Caption)
    End If
    
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
    
    Rem dada
    Rem dada
    Rem dada
    Rem dada

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
    Select Case Val(WEmpresa)
        Case 1
            ZZNumero = "30549165083"
        Case Else
            ZZNumero = "30610524598"
    End Select
    
    ZZNumero = ZZNumero + "01"
    
    Select Case Val(WEmpresa)
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





