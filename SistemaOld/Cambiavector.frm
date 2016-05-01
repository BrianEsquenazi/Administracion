VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgcambiavector 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Gastos de Despacho"
   ClientHeight    =   8370
   ClientLeft      =   225
   ClientTop       =   480
   ClientWidth     =   11490
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   11490
   Visible         =   0   'False
   Begin VB.Frame PantaFecha 
      Height          =   2055
      Left            =   4080
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton CancelaFecha 
         Caption         =   "   Cancela Actualizacion"
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
         Left            =   480
         TabIndex        =   33
         Top             =   1320
         Width           =   1935
      End
      Begin MSMask.MaskEdBox FechaLLegada 
         Height          =   285
         Left            =   600
         TabIndex        =   34
         Top             =   840
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Indique Fecha Prevista de Arribo"
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
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame IngresaObservaciones 
      Caption         =   "Ingreso de Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   2040
      TabIndex        =   46
      Top             =   960
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox WTituloII 
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox WTituloII 
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox WTituloII 
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1080
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3II 
         Height          =   285
         Left            =   3960
         TabIndex        =   52
         Top             =   720
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
      Begin VB.CommandButton Command1 
         Caption         =   "Fin de Ingreso"
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
         Left            =   2760
         TabIndex        =   50
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox WTexto1II 
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
         Left            =   2520
         TabIndex        =   49
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox WCombo1II 
         Height          =   315
         Left            =   2520
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto2II 
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
         Left            =   3360
         TabIndex        =   47
         Top             =   720
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVectorII 
         Height          =   4575
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8070
         _Version        =   327680
         BackColor       =   16777152
      End
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame IngresaDerechos 
      Caption         =   "Ingresos de Derechos de Estadistica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   1080
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   7095
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
         TabIndex        =   41
         Top             =   2400
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2400
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2400
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
         Left            =   2880
         TabIndex        =   38
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   37
         Top             =   2400
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
         Left            =   2040
         TabIndex        =   36
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton FinIngreso 
         Caption         =   "Fin de Ingreso"
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
         Left            =   2760
         TabIndex        =   29
         Top             =   5160
         Width           =   2175
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   3480
         TabIndex        =   42
         Top             =   1920
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
         Height          =   4575
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8070
         _Version        =   327680
         BackColor       =   16777152
      End
   End
   Begin VB.CommandButton IngreFecha 
      Caption         =   "Actualiza Fecha Llegada"
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
      Left            =   5400
      TabIndex        =   31
      Top             =   7080
      Width           =   2055
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
      Left            =   8760
      TabIndex        =   25
      Top             =   480
      Width           =   2055
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
      Left            =   4080
      MaxLength       =   30
      TabIndex        =   24
      Top             =   480
      Width           =   3015
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
      Left            =   1080
      TabIndex        =   19
      Top             =   6480
      Width           =   975
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
      TabIndex        =   17
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9000
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
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
      Left            =   5400
      TabIndex        =   15
      Top             =   6480
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
      Left            =   7800
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4080
      TabIndex        =   13
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
      Left            =   1800
      MaxLength       =   6
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
      Height          =   500
      Left            =   0
      TabIndex        =   10
      Top             =   6480
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
      Left            =   4320
      TabIndex        =   9
      Top             =   6480
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
      Left            =   2160
      TabIndex        =   7
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   6975
      Begin VB.TextBox WImporte 
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
         TabIndex        =   26
         Text            =   " "
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox WConcepto 
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
         MaxLength       =   6
         TabIndex        =   18
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WLinea 
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
         Left            =   0
         TabIndex        =   8
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1815
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
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Concepto"
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
         TabIndex        =   20
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
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3735
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
      Left            =   3240
      TabIndex        =   4
      Top             =   6480
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4215
      Left            =   0
      OleObjectBlob   =   "Cambiavector.frx":0000
      TabIndex        =   3
      Top             =   960
      Width           =   7215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   7680
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
      Height          =   7080
      ItemData        =   "Cambiavector.frx":09EA
      Left            =   7800
      List            =   "Cambiavector.frx":09F1
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Derechos 
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
      Height          =   500
      Left            =   6480
      TabIndex        =   30
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo Costo"
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
      Left            =   7320
      TabIndex        =   45
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   3000
      TabIndex        =   23
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label8 
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
      Left            =   7320
      TabIndex        =   22
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Height          =   375
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
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Prgcambiavector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 3 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WAnterior As Integer
Private Auxiliar(100, 5) As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstGasimpo As Recordset
Dim spGasimpo As String
Dim XParam As String
Dim WProveedor As String
Dim XDerechos As Double
Dim WDerechos As String
Dim Vector(100) As String
Dim WImpoDerechos As String
Dim XImpoDerechos As Double
Dim WCalcula(100, 10) As String
Dim EmpresaAnterior As String
Dim EmpresaOrden As Integer
Dim XOrden As Integer

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametrosII(10, 10) As Double
Dim WFormatoII(10) As String

Dim WBorraII(1000, 10) As String
Dim WParametrosIIII(10, 10) As Double
Dim WFormatoIIII(10) As String

Dim WControl As String

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""

    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""
    WLinea.Text = ""
    
    WConcepto.SetFocus
    
End Sub


Private Sub cmdClose_Click()

    Call Limpia_Click

    PrgMovgas.Hide
    Unload Me
    Menu.Show
    
End Sub


Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Conceptos de Gastos de Importacion"

     Opcion.Visible = True
     
 End Sub

Private Sub Derechos_Click()

    IngresaDerechos.Height = 6015
    IngresaDerechos.Left = 1080
    IngresaDerechos.Top = 960
    IngresaDerechos.Width = 7095

    Call Busca_Empresa
    If EmpresaOrden = 0 Then
        Exit Sub
    End If

    EmpresaAnterior = WEmpresa
    Select Case EmpresaOrden
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
        Case Else
    End Select
    
    Call Limpia_VectorII
    
    Erase Vector
    Entre = 0

    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    Entra = Entra + 1
                    WVector1II.Row = Entra
                    WVector1II.Col = 1
                    WVector1II.Text = rstOrden!Articulo
                    WVector1II.Col = 2
                    WVector1II.Text = ""
                    XDerechos = IIf(IsNull(rstOrden!Derechos), "0", rstOrden!Derechos)
                    WVector1II.Col = 3
                    WVector1II.Text = Str$(XDerechos)
                    WVector1II.Text = Pusing("###,###.##", WVector1II.Text)
                    Vector(Entra) = rstOrden!Clave
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    For Cicla = 1 To 100
        WArticulo = WVector1II.TextMatrix(Cicla, 1)
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1II.Row = Cicla
            WVector1II.Col = 2
            WVector1II.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    Next Cicla
    
    Select Case Val(EmpresaAnterior)
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
        Case Else
    End Select
    
    IngresaDerechos.Visible = True
    
    WVector1II.Col = 3
    WVector1II.Row = 1
    WVector1II.TopRow = 1
    Call StartEditII
    
End Sub

Private Sub FinIngreso_Click()

    EmpresaAnterior = WEmpresa
    Select Case EmpresaOrden
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
        Case Else
    End Select
    
    For Cicla = 1 To 100
        spOrden = "ConsultaOrden " + "'" + Vector(Cicla) + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WClave = Vector(Cicla)
            WDerechos = WVector1II.TextMatrix(Cicla, 3)
            rstOrden.Close
            XParam = "'" + WClave + "','" _
                + WDerechos + "'"
            spOrden = "ModificaOrdenDerechos " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        End If
    Next Cicla
    
    Select Case Val(EmpresaAnterior)
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
        Case Else
    End Select
    
    IngresaDerechos.Visible = False
    
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
            spGasimpo = "ListaGasimpo"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstGasimpo.RecordCount > 0 Then
                With rstGasimpo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstGasimpo!Codigo) + " " + rstGasimpo!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstGasimpo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstGasimpo.Close
            End If
            
        Case Else
    End Select
    Pantalla.Visible = True
            
End Sub

Private Sub DBGrid1_GotFocus()

    DBGrid1.Col = 0
    If Val(DBGrid1.Text) <> 0 Then
        WLinea.Text = DBGrid1.Row + 1
        WConcepto.Text = DBGrid1.Text
            Else
        WConcepto.Text = ""
        WLinea.Text = ""
    End If
    
    DBGrid1.Col = 1
    WDescripcion.Caption = DBGrid1.Text

    DBGrid1.Col = 2
    If Val(DBGrid1.Text) <> 0 Then
        WImporte.Text = DBGrid1.Text
            Else
        WImporte.Text = ""
    End If
    
    WConcepto.SetFocus

End Sub

Private Sub Graba_Click()

    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        WMarca = IIf(IsNull(rstMovgas!Marca), "", rstMovgas!Marca)
        rstMovgas.Close
        If WMarca = "X" Then
            m$ = "Esta carpeta ya fue actualizada"
            A% = MsgBox(m$, 0, "Ingreso de Gastos")
            Exit Sub
        End If
    End If
    
    XPasa = "S"
    Call Busca_Empresa
    
    EmpresaAnterior = WEmpresa
    Select Case EmpresaOrden
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
        Case Else
    End Select

    spOrden = "ListaOrden " + "'" + Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
        If WTipoOrden <> 1 Then
            rstOrden.Close
            m$ = "La Orden no es de importacion"
            A% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
            XPasa = "N"
                Else
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
            Else
        m$ = "Nro. de Orden de Compra Inexistente"
        A% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
        XPasa = "N"
    End If
    
    Select Case Val(EmpresaAnterior)
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
        Case Else
    End Select
    
    If XPasa = "N" Then
        Exit Sub
    End If
    
    If XPasa = "S" Then
    
        EmpresaAnterior = WEmpresa
        Select Case EmpresaOrden
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
            Case Else
        End Select
    
        XImpoDerechos = 0
        Erase WCalcula
        Ingre = 0
    
        spOrden = "ListaOrden " + "'" + Orden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        Ingre = Ingre + 1
                        WCalcula(Ingre, 1) = rstOrden!Articulo
                        WCalcula(Ingre, 2) = Str$(rstOrden!Cantidad)
                        otro = IIf(IsNull(rstOrden!Derechos), 0, rstOrden!Derechos)
                        WCalcula(Ingre, 3) = Str$(otro)
                        WCalcula(Ingre, 4) = Str$(rstOrden!Precio)
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        End If
    
        Rem For Cicla = 1 To Ingre
        Rem     spArticulo = "ConsultaArticulo " + "'" + WCalcula(Cicla, 1) + "'"
        Rem     Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstArticulo.RecordCount > 0 Then
        Rem         WCalcula(Cicla, 4) = Str$(rstArticulo!Flete)
        Rem         rstArticulo.Close
        Rem     End If
        Rem Next Cicla
    
        For Cicla = 1 To Ingre
            XImpoDerechos = XImpoDerechos + ((Val(WCalcula(Cicla, 4)) * Val(WCalcula(Cicla, 2))) * Val(WCalcula(Cicla, 3)) / 100)
        Next Cicla
    
        spOrden = "ListaOrden " + "'" + Orden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
            Origen.SetFocus
                Else
            Orden.SetFocus
        End If
        
        Select Case Val(EmpresaAnterior)
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
            Case Else
        End Select
        
        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                    
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        DBGrid1.Col = 0
        DBGrid1.Text = ""
    
        spMovgas = "BorrarMovgas " + "'" + Carpeta.Text + "'"
        Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenDynaset, dbSQLPassThrough)
    
        Renglon = 0
        DBGrid1.Refresh
        
        For A = 0 To 3
        
            Suma = A * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
                
                WRow = iRow
                DBGrid1.Row = WRow
                    
                DBGrid1.Col = 0
                concepto = DBGrid1.Text
                    
                DBGrid1.Col = 2
                Importe = DBGrid1.Text
                    
                If Val(concepto) <> 0 Then
                        
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(Carpeta.Text)
                    Call Ceros(Auxi1, 6)
                
                    WClave = Auxi1 + Auxi
                    WCarpeta = Carpeta.Text
                    WRenglon = Str$(Renglon)
                    WFecha = Fecha.Text
                    WDerechos = ""
                    WOrden = Orden.Text
                    WConcepto = concepto
                    WImporte = Importe
                    WAuxiliar = ""
                    WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    WMoneda = Str$(Moneda.ListIndex)
                    WOrigen = Origen.Text
                    WMarca = ""
                    WImpoDerechos = Str$(XImpoDerechos)
                    WFechaLlegada = ""
                    WOrdFechaLLegada = ""
                    WCostoFlete = ""
                    WGastos = ""
                    WPagado = ""
                    WEmpreOtro = Str$(EmpresaOrden)
                
                    XParam = "'" + WClave + "','" _
                         + WCarpeta + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WDerechos + "','" _
                         + WOrden + "','" _
                         + WConcepto + "','" _
                         + WImporte + "','" _
                         + WAuxiliar + "','" _
                         + WOrdFecha + "','" _
                         + WProveedor + "','" _
                         + WOrigen + "','" _
                         + WMoneda + "','" _
                         + WMarca + "','" _
                         + WImpoDerechos + "','" _
                         + WFechaLlegada + "','" _
                         + WOrdFechaLLegada + "','" _
                         + WCostoFlete + "','" _
                         + WGastos + "','" _
                         + WPagado + "','" _
                         + WEmpreOtro + "'"
                         
                    spMovgas = "AltaMovgas " + XParam
                    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
            Next iRow
            
        Next A
            
        Call Limpia_Click
        
        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
    End If
    
    Carpeta.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""
    
    WConcepto.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Pantalla.Visible = False
    
    WLinea.Text = ""
    WConcepto.Text = ""
    WDescripcion.Caption = ""
    WImporte.Text = ""

    Carpeta.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Orden.Text = ""
    Rem Derechos.Text = ""
    Origen.Text = ""
    
    Moneda.ListIndex = 0
    
    For A = 0 To 3
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 2
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    Carpeta.Text = ""
    
    spMovgas = "ListaMovgasTotal"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        
        With rstMovgas
            .MoveLast
            Carpeta.Text = rstMovgas!Carpeta + 1
        End With
    
        rstMovgas.Close
        
    End If
    
    DBGrid1.FirstRow = 0
    Renglon = 0

    Carpeta.SetFocus

End Sub

Private Sub WConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spGasimpo = "ConsultaGasimpo " + "'" + WConcepto.Text + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            WDescripcion.Caption = rstGasimpo!Nombre
            WImporte.SetFocus
                Else
            WConcepto.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WImporte_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WImporte.Text = Pusing("###,###.##", WImporte.Text)
        If Leyenda.ListIndex = 2 And Val(WConcepto.Text) = 4 Then
            m$ = "El costo de la mercaderia ya incluye el flete (Cyf)"
            A% = MsgBox(m$, 0, "Ingreso de Gastos de Importacion")
            Exit Sub
        End If
        Call Alta_Vector
        Call Ingresa_Click
        WConcepto.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
        
            spGasimpo = "ConsultaGasimpo " + "'" + Claveven$ + "'"
            Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
            If rstGasimpo.RecordCount > 0 Then
                    WConcepto.Text = rstGasimpo!Codigo
                    WDescripcion.Caption = rstGasimpo!Nombre
                    
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstGasimpo!Codigo
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstGasimpo!Nombre
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    WImporte.SetFocus
                    
            End If

        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3
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

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 2, 0 To 40)

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
For i = 0 To 2
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Concepto"
             DBGrid1.Columns(newcnt).Width = 1000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Importe"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
 DBGrid1.Font.Bold = True
 
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem With rstmovgas
    Rem     .Index = "Clave"
    Rem Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Carpeta.text = !Carpeta + 1
    Rem             Else
    Rem         Carpeta.text = ""
    Rem     End If
    Rem End With
    
    Moneda.Clear
    Moneda.AddItem "Dolar"
    Moneda.AddItem "Euro"
    
    Moneda.ListIndex = 0
    
    Leyenda.Clear
    
    Leyenda.AddItem ""
    Leyenda.AddItem "FOB"
    Leyenda.AddItem "CyF"
    
    Leyenda.ListIndex = 0
    
    Carpeta.Text = ""
    
    spMovgas = "ListaMovgasTotal"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        With rstMovgas
            .MoveLast
            Carpeta.Text = rstMovgas!Carpeta + 1
        End With
        rstMovgas.Close
    End If
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovgas.Caption = "Ingreso de Gastos de Despacho :  " + !Nombre
        End If
    End With
 
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Carpeta.SetFocus
    
End Sub

Private Sub Proceso_Click()

    For A = 0 To 3
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 2
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    Erase Auxiliar
    
    spMovgas = "Listamovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Fecha.Text = rstMovgas!Fecha
                    Orden.Text = rstMovgas!Orden
                    Rem Derechos.Text = rstMovgas!Derechos
                    Rem Derechos.Text = Pusing("###,###.##", Derechos.Text)
                    Origen.Text = rstMovgas!Origen
                    Moneda.ListIndex = rstMovgas!Moneda
                    EmpresaOrden = rstMovgas!Empresa
        
                    Renglon = Renglon + 1
            
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
                
                    DBGrid1.Col = 0
                    DBGrid1.Text = rstMovgas!concepto
                    Auxi1 = rstMovgas!concepto
                
                    DBGrid1.Col = 2
                    DBGrid1.Text = Pusing("###,###.##", rstMovgas!Importe)
                
                    Auxiliar(Renglon, 1) = Auxi1
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovgas.Close
                
    End If
    
    Call Orden_KeyPress(13)
    
    WRenglon = Renglon
    Renglon = 0
    
    For Renglon = 1 To WRenglon
    
        Auxi1 = Auxiliar(Renglon, 1)
    
        Lugar1 = Int((Renglon - 1) / 10) * 10
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
        
        spGasimpo = "ConsultaGAsimpo " + "'" + Auxi1 + "'"
        Set rstGasimpo = db.OpenRecordset(spGasimpo, dbOpenSnapshot, dbSQLPassThrough)
        If rstGasimpo.RecordCount > 0 Then
            DBGrid1.Col = 1
            DBGrid1.Text = rstGasimpo!Nombre
            WConcepto.SetFocus
        End If
        
    Next Renglon
    
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
    
    WConcepto.SetFocus

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
            
            DBGrid1.Text = WConcepto.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
                
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WConcepto.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Caption
                
            DBGrid1.Col = 2
            DBGrid1.Text = Pusing("###,###.##", WImporte.Text)
            
            DBGrid1.Row = Renglon
            DBGrid1.Col = 0
            
    End If

End Sub

Private Sub Carpeta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(Carpeta.Text) <> 0 Then
    
        spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
        Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovgas.RecordCount > 0 Then
        
            Fecha.Text = rstMovgas!Fecha
            rstMovgas.Close
            Call Proceso_Click
            
                Else
                
            WCarpeta = Carpeta.Text
            Call Limpia_Click
            Carpeta.Text = WCarpeta
            Call Busca_Empresa_Carpeta
            Orden.Text = Str$(XOrden)
            Call Orden_KeyPress(13)
            Fecha.SetFocus
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Orden.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Orden.Text <> "" Then
        
            Call Busca_Empresa
            EmpresaAnterior = WEmpresa
            
            Select Case EmpresaOrden
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
                Case Else
            End Select
        
            spOrden = "ListaOrden " + "'" + Orden.Text + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
                If WTipoOrden <> 1 Then
                    rstOrden.Close
                    m$ = "La Orden no es de importacion"
                    A% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                    Orden.SetFocus
                        Else
                    WProveedor = rstOrden!Proveedor
                    Leyenda.ListIndex = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
                    rstOrden.Close
                    Origen.SetFocus
                End If
                    Else
                m$ = "no existe el numero de Orden de Compra"
                A% = MsgBox(m$, 0, "Carga de Gastos de Importacion")
                Orden.SetFocus
            End If
            
            Select Case Val(EmpresaAnterior)
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
                Case Else
            End Select
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Origen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WConcepto.SetFocus
    End If
End Sub

Private Sub Busca_Empresa()

    EmpresaOrden = 0
    EmpresaAnterior = WEmpresa

    For Va = 1 To 8
    
        Select Case Va
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
            Case Else
        End Select
        
        spOrden = "ListaOrden " + "'" + Orden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            If WTipoOrden = 1 Then
                EmpresaOrden = Va
            End If
            rstOrden.Close
        End If
    
    Next Va
    
    Select Case Val(EmpresaAnterior)
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
        Case Else
    End Select
    
End Sub


Private Sub Busca_Empresa_Carpeta()

    XOrden = 0
    EmpresaAnterior = WEmpresa

    For Va = 1 To 8
    
        Select Case Va
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
            Case Else
        End Select
        
        spOrden = "ListaOrdenCarpeta " + "'" + Carpeta.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            If WTipoOrden = 1 Then
                XOrden = rstOrden!Orden
            End If
            rstOrden.Close
        End If
    
    Next Va
    
    Select Case Val(EmpresaAnterior)
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
        Case Else
    End Select
    
End Sub

Private Sub IngreFecha_Click()
    FechaLLegada.Text = "  /  /    "
    PantaFecha.Visible = True
    FechaLLegada.SetFocus
End Sub

Private Sub CancelaFecha_Click()
    PantaFecha.Visible = False
End Sub

Private Sub FechaLLegada_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaLLegada.Text, Auxi)
        If Auxi = "S" Then
            Call Busca_Empresa
            If EmpresaOrden <> 0 Then
                EmpresaAnterior = WEmpresa
                Select Case EmpresaOrden
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
                    Case Else
                End Select
    
                XParam = "'" + Orden.Text + "','" _
                         + FechaLLegada.Text + "','" _
                         + FechaLLegada.Text + "'"
    
                spOrden = "ModificaOrdenFechaLLegada " + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
                Select Case Val(EmpresaAnterior)
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
                    Case Else
                End Select
            End If
            PantaFecha.Visible = False
                Else
            FechaLLegada.SetFocus
        End If
    End If
End Sub



Rem
Rem Controles de la WVector1II
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector1II.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1II.Left = WVector1II.CellLeft + WVector1II.Left
            WTexto1II.Top = WVector1II.CellTop + WVector1II.Top
            WTexto1II.Width = WVector1II.CellWidth
            WTexto1II.Height = WVector1II.CellHeight
            WTexto1II.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1II.Text = WVector1II.Text
                    WTexto1II.SelStart = Len(WTexto1II.Text)
                Case Else
                    WTexto1II.Text = Chr$(KeyAscii)
                    WTexto1II.SelStart = 1
            End Select
            WTexto1II.Visible = True
            WTexto1II.SetFocus
        Case 1
            WTexto2II.Left = WVector1II.CellLeft + WVector1II.Left
            WTexto2II.Top = WVector1II.CellTop + WVector1II.Top
            WTexto2II.Width = WVector1II.CellWidth
            WTexto2II.Height = WVector1II.CellHeight
            WTexto2II.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2II.Text = WVector1II.Text
                    Rem WTexto2II.SelStart = Len(WTexto2II.Text)
                    WTexto2II.SelStart = 0
                Case Else
                    WTexto2II.Text = Chr$(KeyAscii)
                    WTexto2II.SelStart = 1
            End Select
            WTexto2II.Visible = True
            WTexto2II.SetFocus
        Case 2
            WTexto3II.Left = WVector1II.CellLeft + WVector1II.Left
            WTexto3II.Top = WVector1II.CellTop + WVector1II.Top
            WTexto3II.Width = WVector1II.CellWidth
            WTexto3II.Height = WVector1II.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1II.Text) = 10 Then
                        WTexto3II.Text = WVector1II.Text
                            Else
                        WTexto3II.Text = "  /  /    "
                    End If
                    WTexto3II.SelStart = 0
                Case Else
                    WTexto3II.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3II.SelStart = 1
            End Select
            WTexto3II.Visible = True
            WTexto3II.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo1II.Visible Then
        Pasa = 0
        WVector1II.Text = WCombo1II.Text
        WCombo1II.Visible = False
            Else
        If WTexto1II.Visible Then
            Pasa = 1
            WVector1II.Text = WTexto1II.Text
            WTexto1II.Visible = False
                Else
            If WTexto2II.Visible Then
                Pasa = 1
                WVector1II.Text = WTexto2II.Text
                WTexto2II.Visible = False
                    Else
                If WTexto3II.Visible Then
                    Pasa = 1
                    WVector1II.Text = WTexto3II.Text
                    WTexto3II.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector1II.Col) <> "" Then
            WVector1II.Text = Pusing(WFormatoII(WVector1II.Col), WVector1II.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo1II.Left = WVector1II.CellLeft + WVector1II.Left
    WCombo1II.Top = WVector1II.CellTop + WVector1II.Top
    WCombo1II.Width = WVector1II.CellWidth
    WCombo1II.Visible = True
    WCombo1II.SetFocus
End Sub

Private Sub WTexto1II_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1II.Text = ""
            
        Rem F1
        Case 113
            WTexto1II.Text = WVector1II.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1II.SetFocus
            DoEvents
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_WVector1II
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row < WVector1II.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow < WVector1II.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow + 12
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow - 12 > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow - 12
                    WVector1II.Row = WVector1II.TopRow
                        Else
                    WVector1II.TopRow = 1
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto2II_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2II.Text = ""
            
        Rem F1
        Case 113
            WTexto2II.Text = WVector1II.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1II.SetFocus
            DoEvents
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_WVector1II
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row < WVector1II.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow < WVector1II.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow + 12
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow - 12 > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow - 12
                    WVector1II.Row = WVector1II.TopRow
                        Else
                    WVector1II.TopRow = 1
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto3II_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3II.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3II.Text = WVector1II.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1II.SetFocus
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_WVector1II
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row < WVector1II.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.Row > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.Row = WVector1II.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow < WVector1II.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow + 12
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1II.SetFocus
            DoEvents
            If WVector1II.TopRow - 12 > WVector1II.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1II.TopRow = WVector1II.TopRow - 12
                    WVector1II.Row = WVector1II.TopRow
                        Else
                    WVector1II.TopRow = 1
                    WVector1II.Row = WVector1II.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1II_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2II_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3II_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1II_Click()
    WVector1II.SetFocus
End Sub


Private Sub WVector1II_Click()
    StartEditII
End Sub

Private Sub WVector1II_LeaveCell()
    EndEditII
End Sub

Private Sub WVector1II_GotFocus()
    EndEditII
End Sub

Private Sub WVector1II_KeyPress(KeyAscii As Integer)
    XColumna = WVector1II.Col
    Select Case WParametrosII(4, WVector1II.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector1II.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1II.Clear
            WCombo1II.AddItem "Campo1"
            WCombo1II.AddItem "Campo2"
            On Error Resume Next
            WCombo1II.Text = WVector1II.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector1II.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector1II()
    Select Case WVector1II.Col
        Case 3
            If WVector1II.Row < WVector1II.Rows - 1 Then
                WVector1II.Row = WVector1II.Row + 1
            End If
            WVector1II.Col = 3
        Case Else
            If WVector1II.Col < WVector1II.Cols - 1 Then
                WVector1II.Col = WVector1II.Col + 1
            End If
    End Select
    WVector1II.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector1II.Col
    XFila = WVector1II.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            WVector1II.Col = XColumna
        Case Else
            WVector1II.Col = XColumna
    End Select
End Sub

Private Sub Limpia_VectorII()

    WVector1II.Clear

    Rem ponga la WVector1II en negritas
    WVector1II.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1II.FontName = WVector1II.FontName
    WTexto1II.FontSize = WVector1II.FontSize
    WTexto1II.Visible = False
    WTexto2II.FontName = WVector1II.FontName
    WTexto2II.FontSize = WVector1II.FontSize
    WTexto2II.Visible = False
    WTexto3II.FontName = WVector1II.FontName
    WTexto3II.FontSize = WVector1II.FontSize
    WTexto3II.Visible = False
    WCombo1II.FontName = WVector1II.FontName
    WCombo1II.FontSize = WVector1II.FontSize
    WCombo1II.Visible = False

    ' Establesco loa Valores de la WVector1II
    
    WVector1II.FixedCols = 1
    WVector1II.Cols = 4
    WVector1II.FixedRows = 1
    WVector1II.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1II.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1II.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1II.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1II.ColWidth(0) = 200
    WVector1II.Row = 0
    For Ciclo = 1 To WVector1II.Cols - 1
        WVector1II.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1II.Text = "Articulo"
                WVector1II.ColWidth(Ciclo) = 1300
                WVector1II.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 12
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector1II.Text = "Descripcion"
                WVector1II.ColWidth(Ciclo) = 3400
                WVector1II.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector1II.Text = "Derechos"
                WVector1II.ColWidth(Ciclo) = 1300
                WVector1II.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 10
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = "###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1II.Row = 0
    For Ciclo = 1 To WVector1II.Cols - 1
        WVector1II.Col = Ciclo
        WTituloII(Ciclo).Text = WVector1II.Text
        WTituloII(Ciclo).Left = WVector1II.CellLeft + WVector1II.Left
        WTituloII(Ciclo).Top = WVector1II.CellTop + WVector1II.Top
        WTituloII(Ciclo).Width = WVector1II.CellWidth
        WTituloII(Ciclo).Height = WVector1II.CellHeight
        WTituloII(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector1II
    
    WAncho = 340
    For Ciclo = 0 To WVector1II.Cols - 1
        WAncho = WAncho + WVector1II.ColWidth(Ciclo)
    Next Ciclo
    WVector1II.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1II.Font.Name
    Font.Size = WVector1II.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1II.AllowUserResizing = flexResizeBoth
    
    WVector1II.Col = 1
    WVector1II.Row = 1
    
End Sub

Private Sub WVector1II_Scroll()
    WTexto1II.Visible = False
    WTexto2II.Visible = False
    WTexto3II.Visible = False
End Sub

