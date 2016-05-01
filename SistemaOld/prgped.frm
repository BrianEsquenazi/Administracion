VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPed 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Pedidos"
   ClientHeight    =   8595
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   Visible         =   0   'False
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
      Height          =   3015
      Left            =   8880
      TabIndex        =   18
      Top             =   5160
      Width           =   2775
      Begin VB.CommandButton AvisoError 
         Caption         =   "Sistema sin Conexion"
         Height          =   1215
         Left            =   600
         Picture         =   "prgped.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
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
         TabIndex        =   68
         Top             =   2520
         Width           =   975
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
         Left            =   1200
         TabIndex        =   67
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   66
         Top             =   2280
         Width           =   975
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
         Left            =   1200
         TabIndex        =   65
         Top             =   2280
         Width           =   1215
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
         Left            =   1200
         TabIndex        =   47
         Top             =   2040
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
         TabIndex        =   46
         Top             =   2040
         Width           =   975
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
         Left            =   1200
         TabIndex        =   45
         Top             =   1800
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
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   975
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
         Left            =   1200
         TabIndex        =   43
         Top             =   1560
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
         TabIndex        =   42
         Top             =   1560
         Width           =   975
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
         TabIndex        =   29
         Top             =   1320
         Width           =   975
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
         TabIndex        =   28
         Top             =   1080
         Width           =   975
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
         Left            =   1200
         TabIndex        =   27
         Top             =   1320
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
         Left            =   1200
         TabIndex        =   26
         Top             =   1080
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
         Left            =   1200
         TabIndex        =   24
         Top             =   720
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
         Left            =   1200
         TabIndex        =   23
         Top             =   480
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
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible"
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
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame MuestraCosto 
      Height          =   2535
      Left            =   5640
      TabIndex        =   55
      Top             =   5640
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox FechaCotiza 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   64
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton CerrarPanta 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   1320
         TabIndex        =   62
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox CostoReposicion 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   60
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox CostoStd 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   58
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox CostoUltCpa 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   2280
         TabIndex        =   56
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Cot."
         Height          =   375
         Left            =   240
         TabIndex        =   63
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reposicion"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Std."
         Height          =   375
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo Ult. Cpa"
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox FactorPT 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   53
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox CostoPT 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   51
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox FechaPrecio 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   50
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Item"
      Height          =   855
      Left            =   10200
      TabIndex        =   39
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Termi 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   3840
      TabIndex        =   38
      Top             =   5520
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra2 
      Height          =   2535
      Left            =   120
      TabIndex        =   35
      Top             =   5640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      _Version        =   327680
      Rows            =   100
      Cols            =   4
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra1 
      Height          =   2535
      Left            =   5640
      TabIndex        =   34
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4471
      _Version        =   327680
      Rows            =   100
      Cols            =   4
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   3495
      Left            =   120
      TabIndex        =   32
      Top             =   1560
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6165
      _Version        =   327680
      Rows            =   100
      Cols            =   6
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
      Left            =   8160
      TabIndex        =   31
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton CtaCte 
      Caption         =   "Cuenta Corriente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      TabIndex        =   25
      Top             =   2640
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   0
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
      Height          =   855
      Left            =   10200
      TabIndex        =   17
      Top             =   1680
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   16
      Text            =   " "
      Top             =   1200
      Width           =   7935
   End
   Begin VB.TextBox Hora 
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
      MaxLength       =   5
      TabIndex        =   14
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox FecEntrega 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
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
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
      Height          =   375
      Left            =   3840
      TabIndex        =   54
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo"
      Height          =   375
      Left            =   3840
      TabIndex        =   52
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Precio"
      Height          =   375
      Left            =   3840
      TabIndex        =   49
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock por Partidas"
      Height          =   375
      Left            =   5640
      TabIndex        =   41
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pedidos Pendientes"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTO"
      Height          =   375
      Left            =   3840
      TabIndex        =   37
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe Pedido"
      Height          =   615
      Left            =   10200
      TabIndex        =   36
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label14 
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
      TabIndex        =   30
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Hora"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Pago 
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
      Left            =   6720
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "C.Pago"
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
      TabIndex        =   8
      Top             =   840
      Width           =   975
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
Attribute VB_Name = "PrgPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxi As String
Private WImpre(10) As String
Private WEnvase(10) As String
Private WVector(6, 3) As String
Private XEnvase(40, 6) As String
Private XLinea As Single
Private WDirentrega As String
Private WInicio As Integer
Private Auxiliar(100, 3) As String
Private XSaldo As Double
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
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim XParam As String
Dim ClavePedido(100)
Dim Producto As String
Dim Costo As Double
Dim ZTipoCosto As Integer

Private Sub CerrarPanta_Click()
    MuestraCosto.Visible = False
End Sub

Private Sub cmdClose_Click()

    With rstEmpresa
        .Close
    End With
    PrgPed.Hide
    Unload Me
    PrgAutoriza.Show
    
End Sub

Private Sub CostoPT_dblclick()

    If Left$(Termi.Text, 2) = "PT" Then
    
        ZTipoCosto = 2
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoUltCpa.Text = Str$(Costo)
        CostoUltCpa.Text = Pusing("###,###.##", CostoUltCpa.Text)
    
        ZTipoCosto = 1
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoStd.Text = Str$(Costo)
        CostoStd.Text = Pusing("###,###.##", CostoStd.Text)
    
        ZTipoCosto = 3
        Producto = Termi.Text
        Call Calcula_Costo(Producto, Costo)
        CostoReposicion.Text = Str$(Costo)
        CostoReposicion.Text = Pusing("###,###.##", CostoReposicion.Text)
        
        FechaCotiza.Text = ""
        
        MuestraCosto.Visible = True
    
            Else

        ZZArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
        spArticulo = "ConsultaArticulo " + "'" + ZZArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            CostoUltCpa.Text = Str$(rstArticulo!Costo1)
            CostoUltCpa.Text = Pusing("###,###.##", CostoUltCpa.Text)
            CostoStd.Text = Str$(rstArticulo!Costo2)
            CostoStd.Text = Pusing("###,###.##", CostoStd.Text)
            ZCosto4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
            CostoReposicion.Text = Str$(ZCosto4)
            CostoReposicion.Text = Pusing("###,###.##", CostoReposicion.Text)
            MuestraCosto.Visible = True
            rstArticulo.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cotiza"
        ZSql = ZSql + " Where Cotiza.Articulo = " + "'" + ZZArti + "'"
        ZSql = ZSql + " Order by Cotiza"
        spCotiza = ZSql
        Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
        If rstCotiza.RecordCount > 0 Then
            With rstCotiza
                .MoveLast
                FechaCotiza.Text = rstCotiza!Fecha
            End With
            rstCotiza.Close
        End If
        
    End If

End Sub

Private Sub CtaCte_Click()
    PCliente = Cliente.Text
    PTipo = 1
    PrgCC.Show
    PTipo = 0
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
    Muestra.ColWidth(1) = 1800
    Muestra.ColWidth(2) = 4000
    Muestra.ColWidth(3) = 1200
    Muestra.ColWidth(4) = 1200
    Muestra.ColWidth(5) = 1200
    
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
    Muestra.Text = "Precio"
    
    Muestra1.ColWidth(0) = 150
    Muestra1.ColWidth(1) = 600
    Muestra1.ColWidth(2) = 900
    Muestra1.ColWidth(3) = 750
    
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Tipo"
    
    Muestra1.Col = 2
    Muestra1.Text = "Partida"
    
    Muestra1.Col = 3
    Muestra1.Text = "Stock"
    
    Muestra2.ColWidth(0) = 100
    Muestra2.ColWidth(1) = 900
    Muestra2.ColWidth(2) = 800
    Muestra2.ColWidth(3) = 1300
    
    Muestra2.Row = 0
    
    Muestra2.Col = 1
    Muestra2.Text = "Cliente"
    
    Muestra2.Col = 3
    Muestra2.Text = "Fecha"
    
    Muestra2.Col = 2
    Muestra2.Text = "Canti."

    Tipoped.Clear
    
    Tipoped.AddItem "Normal"
    Tipoped.AddItem "a Fecha"
    Tipoped.AddItem "Fecha Limite"
    Tipoped.AddItem "Urgente"
    Tipoped.AddItem "Retira Cliente"
    Tipoped.AddItem "Muestra"
    
    Tipoped.ListIndex = 0

    Pedido.Text = WXPed
    
    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
            Fecha.Text = rstPedido!Fecha
            Cliente.Text = rstPedido!Cliente
            FecEntrega.Text = rstPedido!FecEntrega
            Hora.Text = rstPedido!Hora
            Observaciones.Text = rstPedido!Observaciones
            Tipoped.ListIndex = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
            rstPedido.Close
            
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                DesCliente.Caption = rstCliente!Razon
                Pago.Caption = rstCliente!Pago1
                WDirentrega = rstCliente!DirEntrega
                Rem Observaciones.Text = rstCliente!Observaciones
                rstCliente.Close
                
                spPago = "ConsultaPago " + "'" + Pago.Caption + "'"
                Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstPago.RecordCount > 0 Then
                    DesPago.Caption = rstPago!Nombre
                    rstPago.Close
                End If
            End If
            Call Proceso_Click
                Else
            WPedido = Pedido.Text
            Pedido.Text = WPedido
    End If
    
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
    Muestra.Text = "Precio"
    
    
    Erase Auxiliar
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
                        
                        Muestra.Col = 5
                        Muestra.Text = Pusing("###,###.##", rstPedido!Precio)
                        
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
    
    For DA = 1 To WRenglon
    
        Cliente = Auxiliar(DA, 1)
        Terminado = Auxiliar(DA, 2)
        ZZNombreComercial = Trim(Auxiliar(DA, 3))
        
        If Left$(Terminado, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
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
            
                    Muestra.Col = 4
                    Canti = Val(Muestra.Text)
            
                    Muestra.Col = 5
                    Precio = Val(Muestra.Text)
            
                    Total = Total + (Canti * Precio)
                    
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
            
                    Muestra.Col = 4
                    Canti = Val(Muestra.Text)
            
                    Muestra.Col = 5
                    Precio = Val(Muestra.Text)
            
                    Total = Total + (Canti * Precio)
                    
                    rstPrecios.Close
            
                End If
        End Select
        
    Next DA
    
    Total.Text = Pusing("###,###.##", Str$(Total))
    Muestra.Row = 1

End Sub



Private Sub Muestra_Click()

    If Val(WEmpresa) = 1 Then
        Stock1.Caption = "Stock SI"
        Stock2.Caption = "Stock SII"
        Stock3.Caption = "Stock SIII"
        Stock4.Caption = "Stock SIV"
        Stock5.Caption = "Stock SV"
        Stock6.Caption = "Stock SVI"
        Stock7.Caption = "Stock SVII"
            Else
        Stock1.Caption = "Stock PI"
        Stock2.Caption = "Stock PII"
        Stock3.Caption = "Stock PV"
        Stock4.Caption = "Stock PVI"
        Stock5.Caption = ""
        Stock6.Caption = ""
        Stock7.Caption = ""
    End If

    WStock1.Caption = ""
    WStock2.Caption = ""
    WStock3.Caption = ""
    WStock4.Caption = ""
    WStock5.Caption = ""
    Wstock6.Caption = ""
    Wstock7.Caption = ""

    Muestra.Col = 1
    Termi.Text = Muestra.Text
    XProducto = Termi.Text
    
    Muestra1.Clear
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Tipo"
    
    Muestra1.Col = 2
    Muestra1.Text = "Partida"
    
    Muestra1.Col = 3
    Muestra1.Text = "Stock"
    
    Muestra2.Clear
    Muestra2.Row = 0
    
    Muestra2.Col = 1
    Muestra2.Text = "Cliente"
    
    Muestra2.Col = 3
    Muestra2.Text = "Fecha"
    
    Muestra2.Col = 2
    Muestra2.Text = "Canti."

    Renglon = 0
    XStock = 0
    XPedido = 0
    
    If Left$(XProducto, 2) <> "PT" Then
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
                                    Muestra1.Text = Left$(WArti, 2)
                        
                                    Muestra1.Col = 2
                                    Muestra1.Text = rstLaudo!Laudo
                            
                                    Muestra1.Col = 3
                                    Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
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
            
            
            Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
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
                        
                            If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArti Then
                    
                                WArticulo = rstMovguia!Articulo
                                WCantidad = rstMovguia!Cantidad
                                WFecha = rstMovguia!Fecha
                                WCodigo = rstMovguia!Codigo
                                WMovi = rstMovguia!Movi
                                WDestino = rstMovguia!Destino
                                WTipomov = rstMovguia!Tipomov
                                WSaldo = rstMovguia!Saldo
                                
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                            
                                Muestra1.Col = 1
                                Muestra1.Text = Left$(WArti, 2)
                        
                                WPartiOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                                If Trim(WPartiOri) <> "" Then
                                    WParti = WPartiOri
                                        Else
                                    WParti = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                End If
                                
                                Muestra1.Col = 2
                                Muestra1.Text = WParti
                            
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(rstMovguia!Saldo))
                        
                                XStock = XStock + rstMovguia!Saldo
                                
                            End If
                        End If
                
                        .MoveNext
            
                        If .EOF = True Then
                            Exit Do
                        End If
                                                                            
                    Loop
                    End If
            
                End With
                rstMovguia.Close
            End If
            
            
            
            
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
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 3
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 2
                            Muestra2.Text = Pusing("###,###", Str$(XPed))
                        
                            XPedido = XPedido + XPed
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
            
            
            
            
            Cliente.Text = UCase(Cliente.Text)
            Termi.Text = UCase(Termi.Text)
            ZZArti = Left$(Termi.Text, 3) + Right$(Termi.Text, 7)
    
            WClave = Cliente.Text + ZZArti
    
            spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                FechaPrecio.Text = IIf(IsNull(rstPreciosMp!Fecha), "", rstPreciosMp!Fecha)
                rstPreciosMp.Close
            End If
            
            CostoPT.Text = ""
            FactorPT.Text = ""
            
            spArticulo = "ConsultaArticulo " + "'" + ZZArti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                CostoPT.Text = Str$(rstArticulo!Costo2)
                CostoPT.Text = Pusing("###,###.##", CostoPT.Text)
                If Val(CostoPT.Text) <> 0 Then
                    ZZPrecioVenta = Muestra.TextMatrix(Muestra.Row, 5)
                    FactorPT.Text = Str$(Val(ZZPrecioVenta) / Val(CostoPT.Text))
                    FactorPT.Text = Pusing("###.##", FactorPT.Text)
                End If
                rstArticulo.Close
            End If
            
            WStock1.Caption = Pusing("###,###.##", Str$(XStock))
            WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
            WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
            WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
            WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
            Wstock6.Caption = Pusing("###,###.##", Wstock6.Caption)
            Wstock7.Caption = Pusing("###,###.##", Wstock7.Caption)
            
            Stock.Caption = Pusing("###,###.##", Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(Wstock6.Caption) + Val(Wstock7.Caption)))
            StkPedido.Caption = Pusing("###,###.##", Str$(XPedido))
            Disponible.Caption = Pusing("###,###.##", Str$(Val(Stock.Caption) - XPedido))
            
    
        Case Else
            Rem lee pt
    
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
                        
                            If XSaldo <> 0 Then
                        
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                            
                                Muestra1.Col = 1
                                Muestra1.Text = "PT"
                        
                                Muestra1.Col = 2
                                Muestra1.Text = XHoja
                            
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
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
                
                        If rstMovguia!Marca = "X" Then
                                Else
                        If rstMovguia!Tipo = "T" Then
                
                            XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                            XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            Call Redondeo(XSaldo)
                    
                            If XSaldo <> 0 Then
                                Renglon = Renglon + 1
                                Muestra1.Row = Renglon
                        
                                Muestra1.Col = 1
                                Muestra1.Text = "PT"
                
                                Muestra1.Col = 2
                                Muestra1.Text = XLote
                        
                                Muestra1.Col = 3
                                Muestra1.Text = Pusing("###,###", Str$(XSaldo))
                        
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
            End If
    
    
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
                            Muestra2.Text = rstPedido!Cliente
                
                            Muestra2.Col = 3
                            Muestra2.Text = rstPedido!FecEntrega
                            
                            Muestra2.Col = 2
                            Muestra2.Text = Pusing("###,###", Str$(XPed))
                        
                            XPedido = XPedido + XPed
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
            
            WSalidaError = ""
            On Error GoTo Control_error
            
            If Val(WEmpresa) = 1 Then
            
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                         WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
            
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock5.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        Wstock6.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        Wstock7.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        Else
                        
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock1.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                         WStock2.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
            
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock3.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    spTerminado = "ConsultaTerminado " + "'" + Termi.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock4.Caption = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        rstTerminado.Close
                    End If
                    
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
            End If
            
            On Error GoTo 0
            
            
            
            Cliente.Text = UCase(Cliente.Text)
            Termi.Text = UCase(Termi.Text)
            WClave = Cliente.Text + Termi.Text
    
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                FechaPrecio.Text = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
                rstPrecios.Close
            End If
            
            
            If Left$(Termi.Text, 2) = "PT" Then
            
                ZTipoCosto = 1
                Producto = Termi.Text
                Call Calcula_Costo(Producto, Costo)
                CostoPT.Text = Str$(Costo)
                CostoPT.Text = Pusing("###,###.##", CostoPT.Text)
                
                If Val(CostoPT.Text) <> 0 Then
                    ZZPrecioVenta = Muestra.TextMatrix(Muestra.Row, 5)
                    FactorPT.Text = Str$(Val(ZZPrecioVenta) / Val(CostoPT.Text))
                    FactorPT.Text = Pusing("###.##", FactorPT.Text)
                End If
                
            End If
            
            
            WStock1.Caption = Pusing("###,###.##", WStock1.Caption)
            WStock2.Caption = Pusing("###,###.##", WStock2.Caption)
            WStock3.Caption = Pusing("###,###.##", WStock3.Caption)
            WStock4.Caption = Pusing("###,###.##", WStock4.Caption)
            WStock5.Caption = Pusing("###,###.##", WStock5.Caption)
            Wstock6.Caption = Pusing("###,###.##", Wstock6.Caption)
            Wstock7.Caption = Pusing("###,###.##", Wstock7.Caption)
            
            Stock.Caption = Pusing("###,###.##", Str$(Val(WStock1.Caption) + Val(WStock2.Caption) + Val(WStock3.Caption) + Val(WStock4.Caption) + Val(WStock5.Caption) + Val(Wstock6.Caption) + Val(Wstock7.Caption)))
            StkPedido.Caption = Pusing("###,###.##", Str$(XPedido))
            Disponible.Caption = Pusing("###,###.##", Str$(Val(Stock.Caption) - XPedido))
        
    End Select
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Stock1.Visible = False
    Stock2.Visible = False
    Stock3.Visible = False
    Stock4.Visible = False
    Stock5.Visible = False
    Stock6.Visible = False
    Stock7.Visible = False
    WStock1.Visible = False
    WStock2.Visible = False
    WStock3.Visible = False
    WStock4.Visible = False
    WStock5.Visible = False
    Wstock6.Visible = False
    Wstock7.Visible = False
    Label4.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Disponible.Visible = False
    Stock.Visible = False
    StkPedido.Visible = False
    Resume Next
    
End Sub
    
Private Sub Borra_Click()

    Rem Muestra.Col = 4
    Rem Muestra.Text = "0.00"
    
    WClavePedido = ClavePedido(Muestra.Row)
    Articulo = ""
            
    XParam = "'" + Left$(WClavePedido, 6) + "','" _
                + Right$(WClavePedido, 2) + "'"
    spPedido = "ConsultaPedido2 " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        Articulo = rstPedido!Terminado
        Cantidad = rstPedido!Cantidad - rstPedido!Facturado
        WFacturado = Str$(rstPedido!Facturado)
        WClavePedido = rstPedido!Clave
        rstPedido.Close
        XParam = "'" + WClavePedido + "','" _
                    + WFacturado + "'"
        spPedido = "ModificaPedidoCantidad " + XParam
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Select Case Left$(Articulo, 2)
        Case "PT"
            spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                WPedido = Str$(rstTerminado!Pedido - Cantidad)
                WDate = Date$
                rstTerminado.Close
                XParam = "'" + WCodigo + "','" _
                             + WPedido + "','" _
                             + WDate + "'"
                spTerminado = "ModificaTerminadoPedido " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
        Case Else
            Arti = Left$(Articulo, 3) + Right$(Articulo, 7)
            spArticulo = "ConsultaArticulo " + "'" + Arti + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = rstArticulo!Codigo
                WVenta = Str$(rstArticulo!Venta - Cantidad)
                WDate = Date$
                XParam = "'" + WCodigo + "','" _
                             + WVenta + "','" _
                             + WDate + "'"
                rstArticulo.Close
                spArticulo = "ModificaArticuloVenta " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
                    
    End Select
    
    BajaImpre = "N"
    WPedido = Left$(WClavePedido, 6)
        
    spPedido = "ConsultaPedido1 " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
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
        
    If BajaImpre = "N" Then
        spPedido = "ModificaPedidoMarca " + "'" + WPedido + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    End If
    
    Call Proceso_Click
    
End Sub

Private Sub Termi_Click()
Stop
End Sub


Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim ZZVector(100, 2) As String
    Dim ZZAuxiliar(100, 3) As String
    
    Erase ZZAuxiliar
    ZZRenglon = 0
    
    ZZVector(1, 1) = Producto
    ZZVector(1, 2) = "1"
    ZZLugar = 1
    ZZCicla = 0
    
    Costo = 0
    
    Do
        ZZCicla = ZZCicla + 1
        If ZZVector(ZZCicla, 1) <> "" Then
    
            ZZEntra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + ZZVector(ZZCicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            If rstComposicion.RecordCount > 0 Then
                With rstComposicion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            ZZEntra = "N"
                        
                            ZZTipo = rstComposicion!Tipo
                            ZZArticulo1 = rstComposicion!Articulo1
                            ZZArticulo2 = rstComposicion!Articulo2
                            ZZCantidad = rstComposicion!Cantidad
                            
                            Select Case ZZTipo
                                Case "T"
                                    If Producto <> ZZArticulo2 Then
                                        ZZLugar = ZZLugar + 1
                                        ZZVector(ZZLugar, 1) = ZZArticulo2
                                        ZZVector(ZZLugar, 2) = Str$(ZZCantidad * Val(ZZVector(ZZCicla, 2)))
                                    End If
                                Case "M"
                                    ZZRenglon = ZZRenglon + 1
                                    ZZAuxiliar(ZZRenglon, 1) = ZZArticulo1
                                    ZZAuxiliar(ZZRenglon, 2) = ZZCantidad
                                    ZZAuxiliar(ZZRenglon, 3) = ZZVector(ZZCicla, 2)
                                Case Else
                            End Select
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To ZZRenglon
        ZZArticulo = ZZAuxiliar(DA, 1)
        ZZCantidad = ZZAuxiliar(DA, 2)
        ZZCantidadII = ZZAuxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Select Case ZTipoCosto
                Case 1
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
                Case 2
                    WCosto = (ZZCantidad * rstArticulo!Costo1 * Val(ZZCantidadII))
                Case 3
                    Costo4 = IIf(IsNull(rstArticulo!Costo4), "0", rstArticulo!Costo4)
                    If Costo4 = 0 Then
                        Costo4 = IIf(IsNull(rstArticulo!Costo2), "0", rstArticulo!Costo2)
                    End If
                    WCosto = (ZZCantidad * Costo4 * Val(ZZCantidadII))
                Case Else
                    WCosto = (ZZCantidad * rstArticulo!Costo2 * Val(ZZCantidadII))
            End Select
            Costo = Costo + WCosto
            rstArticulo.Close
        End If
    Next DA
    
    
End Sub


