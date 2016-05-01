VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPrecioOtro 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Precios por Cliente"
   ClientHeight    =   7425
   ClientLeft      =   1140
   ClientTop       =   780
   ClientWidth     =   9870
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   9870
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
      Left            =   2400
      TabIndex        =   50
      Top             =   4080
      Visible         =   0   'False
      Width           =   7335
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
      Left            =   3120
      TabIndex        =   48
      Top             =   2160
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   390
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
      Left            =   2520
      TabIndex        =   46
      Top             =   2160
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2760
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   2760
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2760
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Pago 
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
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   40
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Precio 
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
      Left            =   1680
      TabIndex        =   36
      Text            =   " "
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   3360
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   5295
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2760
         TabIndex        =   34
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Acepta"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela"
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
         Left            =   3720
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin MSMask.MaskEdBox HastaTerminado 
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox DesdeTerminado 
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin VB.TextBox HastaCliente 
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
         TabIndex        =   29
         Text            =   " "
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox DesdeCliente 
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
         TabIndex        =   28
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cliente"
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
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta Prodcuto"
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
         TabIndex        =   27
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Desde Producto"
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
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cliente"
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
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wprecios.rpt"
      WindowTitle     =   "Listado de Precios por Cliente"
   End
   Begin VB.TextBox Descripcion 
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   1200
      Width           =   4815
   End
   Begin VB.CommandButton CmdLimpiar 
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
      Height          =   300
      Left            =   1200
      TabIndex        =   20
      Top             =   4800
      Width           =   975
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8760
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton lista 
      Caption         =   "Listado"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   4080
      Width           =   975
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
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
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
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
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
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
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
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
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
      TabIndex        =   4
      Top             =   4800
      Width           =   975
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
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox Opcion 
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
      Left            =   2760
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
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
      Height          =   2700
      ItemData        =   "PreciosOtro.frx":0000
      Left            =   2400
      List            =   "PreciosOtro.frx":0007
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   7335
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3720
      TabIndex        =   49
      Top             =   2160
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   1935
      Left            =   1920
      TabIndex        =   51
      Top             =   2040
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3413
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label DesPago 
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
      Left            =   6600
      TabIndex        =   41
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Condicion de Pago"
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
      Left            =   3480
      TabIndex        =   39
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Fecha 
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
      Left            =   1680
      TabIndex        =   38
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha Mod."
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
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label DesTerminado 
      BackColor       =   &H00FFFF00&
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
      Left            =   3240
      TabIndex        =   18
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
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
      Left            =   3240
      TabIndex        =   17
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Articulos"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblLabels 
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgPrecioOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPago As Recordset
Dim spPago As String

Dim XParam As String

Private WFecha1 As String
Private WFactura1 As String
Private WPrecio1 As String
Private WCantidad1 As String
Private WFecha2 As String
Private WFactura2 As String
Private WPrecio2 As String
Private WCantidad2 As String
Private WFecha3 As String
Private WFactura3 As String
Private WPrecio3 As String
Private WCantidad3 As String
Private WFecha4 As String
Private WFactura4 As String
Private WPrecio4 As String
Private WCantidad4 As String
Private WFecha5 As String
Private WFactura5 As String
Private WPrecio5 As String
Private WCantidad5 As String

Private dada As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Acepta_Click()

    DesdeTerminado.Text = UCase(DesdeTerminado.Text)
    HastaTerminado.Text = UCase(HastaTerminado.Text)
    DesdeCliente.Text = UCase(DesdeCliente.Text)
    HastaCliente.Text = UCase(HastaCliente.Text)
    
    Listado.WindowTitle = "Listado de Precios de Productos Terminados por Cliente"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Uno = "{Precios.Terminado} in " + Chr$(34) + DesdeTerminado.Text + Chr$(34) + " to " + Chr$(34) + HastaTerminado.Text + Chr$(34)
    Dos = " and " + "{Precios.Cliente} in " + Chr$(34) + DesdeCliente.Text + Chr$(34) + " to " + Chr$(34) + HastaCliente.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Precios.Cliente, Precios.Terminado, Precios.Precio, Cliente.Razon, Terminado.Descripcion " _
                        + "From " + DSQ + ".dbo.Precios Precios, " _
                        + DSQ + ".dbo.Cliente Cliente, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where Precios.Cliente = Cliente.Cliente AND Precios.Terminado = Terminado.Codigo AND Precios.Cliente >= '" + DesdeCliente.Text + "' AND Precios.Cliente <= '" + HastaCliente.Text + "' AND Precios.Terminado >= '" + DesdeTerminado.Text + "' AND Precios.Terminado <= '" + HastaTerminado.Text + "'"
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Cliente.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Sub Imprime_Descripcion()

    Rem lee Cliente

    WCliente = Cliente.Text
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
    Rem lee Terminado
    
    WTerminado = Terminado.Text
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = rstTerminado!Descripcion
        rstTerminado.Close
            Else
        DesTerminado.Caption = ""
    End If
    
    WPago = Pago.Text
    spPago = "ConsultaPago " + "'" + Pago.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        DesPago.Caption = rstPago!Nombre
        rstPago.Close
            Else
        DesPago.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Precio.Text) = 0 Then
        Precio.Text = "0"
    End If
    If Val(Pago.Text) = 0 Then
        Pago.Text = "0"
    End If
End Sub

Sub Format_datos()
    Precio.Text = Pusing("###,###.##", Precio.Text)
End Sub

Sub Imprime_Datos()

    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Cliente.Text = rstPrecios!Cliente
        Terminado.Text = rstPrecios!Terminado
        Precio.Text = rstPrecios!Precio
        Descripcion.Text = rstPrecios!Descripcion
        Fecha.Caption = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
        Pago.Text = IIf(IsNull(rstPrecios!Pago), "0", rstPrecios!Pago)
        Call Format_datos
            
        'columna 1
        
        Call Limpia_Vector
                    
        WVector1.Row = 1
    
        If rstPrecios!Cantidad1 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha1
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura1
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio1)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad1)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", dada)
        End If
    
        'columna 2
    
        WVector1.Row = 2
            
        If rstPrecios!Cantidad2 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!fecha2
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura2
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio2)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad2)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", dada)
        End If
    
        'columna 3
    
        WVector1.Row = 3
            
        If rstPrecios!Cantidad3 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha3
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura3
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio3)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad3)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", dada)
        End If
    
        'columna 4
    
        WVector1.Row = 4
            
        If rstPrecios!Cantidad4 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha4
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura4
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio4)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad4)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", dada)
        End If
    
        'columna 5
    
        WVector1.Row = 5
        
        If rstPrecios!Cantidad5 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha5
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura5
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio5)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad5)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", dada)
        End If
        rstPrecios.Close

    End If
    Call Imprime_Descripcion
    
End Sub

Private Sub cmdAdd_Click()
    If Cliente.Text <> "" And Terminado.Text <> "" Then
    
        Cliente.Text = UCase(Cliente.Text)
        Terminado.Text = UCase(Terminado.Text)
    
        WCliente = Cliente.Text
        WTerminado = Terminado.Text
        WClave = Cliente.Text + Terminado.Text
        
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            WPasa = "S"
                Else
            WPasa = "N"
        End If
        
        Call Verifica_datos
        
        WVector1.Row = 1
        WVector1.Col = 4
        Auxi = Val(WVector1.Text)
    
        If Auxi <> 0 Then
            WVector1.Col = 1
            WFecha1 = WVector1.Text
            WVector1.Col = 2
            WFactura1 = WVector1.Text
            WVector1.Col = 3
            WPrecio1 = WVector1.Text
            WVector1.Col = 4
            WCantidad1 = WVector1.Text
                Else
            WFecha1 = ""
            WFactura1 = ""
            WPrecio1 = ""
            WCantidad1 = ""
        End If
        
        WVector1.Row = 2
        WVector1.Col = 4
        Auxi = Val(WVector1.Text)
    
        If Auxi <> 0 Then
            WVector1.Col = 1
            WFecha2 = WVector1.Text
            WVector1.Col = 2
            WFactura2 = WVector1.Text
            WVector1.Col = 3
            WPrecio2 = WVector1.Text
            WVector1.Col = 4
            WCantidad2 = WVector1.Text
                Else
            WFecha2 = ""
            WFactura2 = ""
            WPrecio2 = ""
            WCantidad2 = ""
        End If
        
        WVector1.Row = 3
        WVector1.Col = 4
        Auxi = Val(WVector1.Text)
    
        If Auxi <> 0 Then
            WVector1.Col = 1
            WFecha3 = WVector1.Text
            WVector1.Col = 2
            WFactura3 = WVector1.Text
            WVector1.Col = 3
            WPrecio3 = WVector1.Text
            WVector1.Col = 4
            WCantidad3 = WVector1.Text
                Else
            WFecha3 = ""
            WFactura3 = ""
            WPrecio3 = ""
            WCantidad3 = ""
        End If
        
        WVector1.Row = 4
        WVector1.Col = 4
        Auxi = Val(WVector1.Text)
    
        If Auxi <> 0 Then
            WVector1.Col = 1
            WFecha4 = WVector1.Text
            WVector1.Col = 2
            WFactura4 = WVector1.Text
            WVector1.Col = 3
            WPrecio4 = WVector1.Text
            WVector1.Col = 4
            WCantidad4 = WVector1.Text
                Else
            WFecha4 = ""
            WFactura4 = ""
            WPrecio4 = ""
            WCantidad4 = ""
        End If
        
        WVector1.Row = 5
        WVector1.Col = 4
        Auxi = Val(WVector1.Text)
    
        If Auxi <> 0 Then
            WVector1.Col = 1
            WFecha5 = WVector1.Text
            WVector1.Col = 2
            WFactura5 = WVector1.Text
            WVector1.Col = 3
            WPrecio5 = WVector1.Text
            WVector1.Col = 4
            WCantidad5 = WVector1.Text
                Else
            WFecha5 = ""
            WFactura5 = ""
            WPrecio5 = ""
            WCantidad5 = ""
        End If
        Fecha.Caption = Date$
        
        If WPasa = "N" Then
            XParam = "'" + WClave + "','" + Cliente.Text + "','" + Terminado.Text + "','" + Precio.Text + "','" _
                         + Descripcion.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha.Caption + "','" + Pago.Text + "'"
            Set rstPrecios = db.OpenRecordset("AltaPrecios1 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + WClave + "','" + Cliente.Text + "','" + Terminado.Text + "','" + Precio.Text + "','" _
                         + Descripcion.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha.Caption + "','" + Pago.Text + "'"
            Set rstPrecios = db.OpenRecordset("ModificaPrecios2 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Cliente.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Cliente.Text <> "" And Terminado.Text <> "" Then
    
        Cliente.Text = UCase(Cliente.Text)
        Terminado.Text = UCase(Terminado.Text)
    
        WCliente = Cliente.Text
        WTerminado = Terminado.Text
        WClave = Cliente.Text + Terminado.Text
        
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            T$ = "Precios de Producto Terminado por Cliente"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spPrecios = "BorrarPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Cliente.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Precio.Text = ""
    Descripcion.Text = ""
    DesCliente.Caption = ""
    DesTerminado.Caption = ""
    Fecha.Caption = ""
    Pago.Text = ""
    DesPago.Caption = ""
    
    Call Limpia_Vector
    Cliente.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Cliente.SetFocus
    PrgPrecio.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            Call Imprime_Datos
            Terminado.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Command1_Click()
    Stop
    WClave = "G00065"
    spPrecios = "MOdificaPreciosXX " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenDynaset, dbSQLPassThrough)
    Stop
End Sub



Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            Call Imprime_Datos
            Precio.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
End Sub

Private Sub Precio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Pago.SetFocus
    End If
    ''Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Pago.Text) <> 0 Then
            spPago = "ConsultaPago " + "'" + Pago.Text + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                DesPago.Caption = rstPago!Nombre
                rstPago.Close
                Precio.SetFocus
            End If
                Else
            DesPago.Caption = ""
            Precio.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.Text = UCase(DesdeCliente.Text)
        HastaCliente.Text = DesdeCliente.Text
        HastaCliente.SetFocus
    End If
End Sub

Private Sub HastaCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCliente.Text = UCase(HastaCliente.Text)
        DesdeTerminado.SetFocus
    End If
End Sub

Private Sub DesdeTerminado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeTerminado.Text = UCase(DesdeTerminado.Text)
        HastaTerminado.SetFocus
    End If
End Sub

Private Sub HastaTerminado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaTerminado.Text = UCase(HastaTerminado.Text)
        DesdeCliente.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

     Opcion.Clear
     
     Opcion.AddItem "Precios de Producto Terminado por Cliente"
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Condiciones de Pago"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    'XIndice = 0
    
    Select Case XIndice
        Case 0
            spPrecios = "ListaPrecios"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                With rstPrecios
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstPrecios!Cliente = Cliente.Text Then
                                IngresaItem = rstPrecios!Cliente + " " + rstPrecios!Terminado
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstPrecios!Clave
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
            
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
            
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
        
        Case 2
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTerminado!Codigo
                            WIndice.AddItem IngresaItem
                                .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 3
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
            
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstPago!Pago) + " " + rstPago!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstPago!Pago
                            WIndice.AddItem IngresaItem
                                .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPago.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WClave = WIndice.List(Indice)
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Cliente.Text = rstPrecios!Cliente
                Terminado.Text = rstPrecios!Terminado
                rstPrecios.Close
                Call Imprime_Datos
                    Else
                CmdLimpiar_Click
            End If
            Precio.SetFocus
        
        Case 1
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                rstCliente.Close
                Call Imprime_Datos
                    Else
                Cliente.Text = WCliente
            End If
            Cliente.SetFocus
            
        Case 2
            Indice = Pantalla.ListIndex
            WTerminado = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Terminado.Text = rstTerminado!Codigo
                rstTerminado.Close
                Call Imprime_Datos
                    Else
                Terminado.Text = WTerminado
            End If
            Terminado.SetFocus
            
        Case 3
            Indice = Pantalla.ListIndex
            WPago = WIndice.List(Indice)
            spPago = "ConsultaPago " + "'" + WPago + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                Pago.Text = rstPago!Pago
                rstPago.Close
                Call Imprime_Descripcion
                    Else
                Pago.Text = WPago
            End If
            Pago.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Lista_Click()
    DesdeCliente.Text = ""
    HastaCliente.Text = ""
    DesdeTerminado.Text = "  -     -   "
    HastaTerminado.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    DesdeCliente.SetFocus
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    
    Cliente.Text = ""
    Precio.Text = ""
    Descripcion.Text = ""
    DesCliente.Caption = ""
    Fecha.Caption = ""
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Cliente.Text = rstPrecios!Cliente
            Terminado.Text = rstPrecios!Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveLast
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Precios de Prodcuto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "AnteriorPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveLast
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Precio.SetFocus
    
End Sub

Private Sub Siguiente_Click()

    On Error GoTo WError
    
    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "PosteriorPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
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

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Factura"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Precio Unitario"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 340
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

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPrecio.Caption = "Ingreso de Precios por Cliente :  " + !Nombre
        End If
    End With
End Sub



