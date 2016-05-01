VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEvaluaTransporte 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Evaluacion de Transportistas"
   ClientHeight    =   8250
   ClientLeft      =   285
   ClientTop       =   300
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11430
   Begin VB.CommandButton Baja 
      Caption         =   "Inhabilitacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   71
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox Promedio33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   65
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Promedio22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   64
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Promedio11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   63
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Chofer2 
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
      Left            =   5880
      MaxLength       =   4
      TabIndex        =   62
      Text            =   " "
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Chofer3 
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
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   61
      Text            =   " "
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Chofer1 
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
      TabIndex        =   60
      Text            =   " "
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Camion3 
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
      Left            =   7680
      MaxLength       =   4
      TabIndex        =   53
      Text            =   " "
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Camion2 
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
      Left            =   5880
      MaxLength       =   4
      TabIndex        =   52
      Text            =   " "
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Camion1 
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
      TabIndex        =   51
      Text            =   " "
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Evaluador 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   49
      Top             =   480
      Width           =   5775
   End
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   8520
      TabIndex        =   45
      Top             =   4200
      Width           =   2775
      Begin VB.TextBox Promedio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   10
         TabIndex        =   50
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Criterio de Aceptación > 7"
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
         Height          =   495
         Index           =   10
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Caption         =   "Puntaje Proveedor"
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
         Index           =   9
         Left            =   480
         TabIndex        =   46
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox Promedio5 
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   43
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Promedio4 
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   42
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Promedio3 
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   41
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Promedio2 
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   40
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Promedio1 
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
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   39
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Punto35 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   37
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Punto34 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   36
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Punto33 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   35
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Punto32 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   34
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Punto31 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   33
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Punto25 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   31
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Punto24 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   30
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Punto23 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   29
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Punto22 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   28
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Punto21 
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
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Punto15 
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
      TabIndex        =   23
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Punto14 
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
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Punto13 
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
      TabIndex        =   21
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Punto12 
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
      TabIndex        =   20
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Punto11 
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
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Ano 
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
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   18
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Mes 
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
      MaxLength       =   4
      TabIndex        =   16
      Text            =   " "
      Top             =   480
      Width           =   615
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
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   13
      Text            =   " "
      Top             =   120
      Width           =   1455
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   4320
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "evaluatransporte.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "evaluatransporte.frx":030A
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "evaluatransporte.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "evaluatransporte.frx":0A56
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "evaluatransporte.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "evaluatransporte.frx":11A2
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "evaluatransporte.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "evaluatransporte.frx":18EE
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   8175
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Camiones.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Efluentes de Lavado"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
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
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   66
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
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   1440
      TabIndex        =   69
      Top             =   1200
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
   Begin VB.TextBox Observaciones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   70
      Top             =   5280
      Width           =   8295
   End
   Begin VB.Label Cartel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor Inhabiloitado"
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
      Left            =   8280
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblLabels 
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
      Index           =   16
      Left            =   120
      TabIndex        =   68
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabels 
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
      Index           =   15
      Left            =   120
      TabIndex        =   67
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Dominio2 
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
      Left            =   6480
      TabIndex        =   59
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Dominio3 
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
      Left            =   8280
      TabIndex        =   58
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Dominio1 
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
      Left            =   4680
      TabIndex        =   57
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label DesChofer2 
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
      Left            =   6480
      TabIndex        =   56
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label DesChofer3 
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
      Left            =   8280
      TabIndex        =   55
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label DesChofer1 
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
      Left            =   4680
      TabIndex        =   54
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Evaluador"
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
      Index           =   12
      Left            =   3840
      TabIndex        =   48
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Promedio"
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
      Index           =   20
      Left            =   9720
      TabIndex        =   44
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puntaje"
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
      Index           =   17
      Left            =   7680
      TabIndex        =   38
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puntaje"
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
      Index           =   14
      Left            =   5880
      TabIndex        =   32
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dominio"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   26
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chofer"
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
      Left            =   3000
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puntaje"
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
      Index           =   11
      Left            =   4080
      TabIndex        =   24
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Caption         =   "Periodo"
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
      TabIndex        =   17
      Top             =   480
      Width           =   1455
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
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
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
      Left            =   3840
      TabIndex        =   14
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Recuperación de Envases/ Entrega de Certificados /  Remitos conformados/ otros"
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
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quejas del Cliente"
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
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lblLabels 
      Caption         =   "Estado de Vehiculos"
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
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblLabels 
      Caption         =   "Puntualidad"
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
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Caption         =   "Actualización  de Habilitaciones y Licencias"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3600
      MouseIcon       =   "evaluatransporte.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":203A
      ToolTipText     =   "Impresion "
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1920
      MouseIcon       =   "evaluatransporte.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":2B86
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   240
      MouseIcon       =   "evaluatransporte.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":36D2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1080
      MouseIcon       =   "evaluatransporte.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":421E
      ToolTipText     =   "Elimina el Registro"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4440
      MouseIcon       =   "evaluatransporte.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":4D6A
      ToolTipText     =   "Salida"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2760
      MouseIcon       =   "evaluatransporte.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "evaluatransporte.frx":58B6
      ToolTipText     =   "Consulta de Datos"
      Top             =   4440
      Width           =   480
   End
End
Attribute VB_Name = "PrgEvaluaTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCamion As Recordset
Dim spCamion As String
Dim rstChofer As Recordset
Dim spChofer As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim rstEvaluaI As Recordset
Dim spEvaluaI As String

Dim ZProveedor As String
Dim ZMes As String
Dim ZAno As String
Dim ZClave As String

Dim ZSuma As Double
Dim ZSuma1 As Double
Dim ZSuma2 As Double
Dim ZSuma3 As Double
Dim ZSuma4 As Double
Dim ZSuma5 As Double

Dim ZPromedio As Double
Dim ZPromedio1 As Double
Dim ZPromedio2 As Double
Dim ZPromedio3 As Double
Dim ZPromedio4 As Double
Dim ZPromedio5 As Double

Dim ZPromedio11 As Double
Dim ZPromedio22 As Double
Dim ZPromedio33 As Double

Sub Imprime_Datos()

    ZProveedor = Proveedor.Text
    ZMes = Mes.Text
    ZAno = Ano.Text
    
    Call Ceros(ZProveedor, 11)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZClave = ZProveedor + ZMes + ZAno

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 1"
    
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
    
        Evaluador.Text = Trim(rstEvaluaI!Evaluador)
        Observaciones.Text = IIf(IsNull(rstEvaluaI!Observaciones), "1", rstEvaluaI!Observaciones)
        
        Fecha.Text = rstEvaluaI!Fecha
        Vencimiento.Text = rstEvaluaI!Vencimiento
        
        Camion1.Text = rstEvaluaI!Camion1
        Camion2.Text = rstEvaluaI!Camion2
        Camion3.Text = rstEvaluaI!Camion3
        
        Chofer1.Text = rstEvaluaI!Chofer1
        Chofer2.Text = rstEvaluaI!Chofer2
        Chofer3.Text = rstEvaluaI!Chofer3
        
        Punto11.Text = Str$(rstEvaluaI!Punto11)
        Punto12.Text = Str$(rstEvaluaI!Punto12)
        Punto13.Text = Str$(rstEvaluaI!Punto13)
        Punto14.Text = Str$(rstEvaluaI!Punto14)
        Punto15.Text = Str$(rstEvaluaI!Punto15)
        
        Punto21.Text = Str$(rstEvaluaI!Punto21)
        Punto22.Text = Str$(rstEvaluaI!Punto22)
        Punto23.Text = Str$(rstEvaluaI!Punto23)
        Punto24.Text = Str$(rstEvaluaI!Punto24)
        Punto25.Text = Str$(rstEvaluaI!Punto25)
        
        Punto31.Text = Str$(rstEvaluaI!Punto31)
        Punto32.Text = Str$(rstEvaluaI!Punto32)
        Punto33.Text = Str$(rstEvaluaI!Punto33)
        Punto34.Text = Str$(rstEvaluaI!Punto34)
        Punto35.Text = Str$(rstEvaluaI!Punto35)
        
        Promedio.Text = Str$(rstEvaluaI!Promedio)
        
        Promedio1.Text = Str$(rstEvaluaI!Promedio1)
        Promedio2.Text = Str$(rstEvaluaI!Promedio2)
        Promedio3.Text = Str$(rstEvaluaI!Promedio3)
        Promedio4.Text = Str$(rstEvaluaI!Promedio4)
        Promedio5.Text = Str$(rstEvaluaI!Promedio5)
        
        Promedio11.Text = Str$(rstEvaluaI!Promedio11)
        Promedio22.Text = Str$(rstEvaluaI!Promedio22)
        Promedio33.Text = Str$(rstEvaluaI!Promedio33)
        
        Promedio.Text = Str$(rstEvaluaI!Promedio)
        
        If Val(Punto11.Text) <> 0 Then
            Punto11.Text = Pusing("###,###.##", Punto11.Text)
                Else
            Punto11.Text = ""
        End If
        If Val(Punto12.Text) <> 0 Then
            Punto12.Text = Pusing("###,###.##", Punto12.Text)
                Else
            Punto12.Text = ""
        End If
        If Val(Punto13.Text) <> 0 Then
            Punto13.Text = Pusing("###,###.##", Punto13.Text)
                Else
            Punto13.Text = ""
        End If
        If Val(Punto14.Text) <> 0 Then
            Punto14.Text = Pusing("###,###.##", Punto14.Text)
                Else
            Punto14.Text = ""
        End If
        If Val(Punto15.Text) <> 0 Then
            Punto15.Text = Pusing("###,###.##", Punto15.Text)
                Else
            Punto15.Text = ""
        End If
        
        If Val(Punto21.Text) <> 0 Then
            Punto21.Text = Pusing("###,###.##", Punto21.Text)
                Else
            Punto21.Text = ""
        End If
        If Val(Punto22.Text) <> 0 Then
            Punto22.Text = Pusing("###,###.##", Punto22.Text)
                Else
            Punto22.Text = ""
        End If
        If Val(Punto23.Text) <> 0 Then
            Punto23.Text = Pusing("###,###.##", Punto23.Text)
                Else
            Punto23.Text = ""
        End If
        If Val(Punto24.Text) <> 0 Then
            Punto24.Text = Pusing("###,###.##", Punto24.Text)
                Else
            Punto24.Text = ""
        End If
        If Val(Punto25.Text) <> 0 Then
            Punto25.Text = Pusing("###,###.##", Punto25.Text)
                Else
            Punto25.Text = ""
        End If
        
        If Val(Punto31.Text) <> 0 Then
            Punto31.Text = Pusing("###,###.##", Punto31.Text)
                Else
            Punto31.Text = ""
        End If
        If Val(Punto32.Text) <> 0 Then
            Punto32.Text = Pusing("###,###.##", Punto32.Text)
                Else
            Punto32.Text = ""
        End If
        If Val(Punto33.Text) <> 0 Then
            Punto33.Text = Pusing("###,###.##", Punto33.Text)
                Else
            Punto33.Text = ""
        End If
        If Val(Punto34.Text) <> 0 Then
            Punto34.Text = Pusing("###,###.##", Punto34.Text)
                Else
            Punto34.Text = ""
        End If
        If Val(Punto35.Text) <> 0 Then
            Punto35.Text = Pusing("###,###.##", Punto35.Text)
                Else
            Punto35.Text = ""
        End If
        
        Promedio1.Text = Pusing("###,###.##", Promedio1.Text)
        Promedio2.Text = Pusing("###,###.##", Promedio2.Text)
        Promedio3.Text = Pusing("###,###.##", Promedio3.Text)
        Promedio4.Text = Pusing("###,###.##", Promedio4.Text)
        Promedio5.Text = Pusing("###,###.##", Promedio5.Text)
        
        Promedio11.Text = Pusing("###,###.##", Promedio11.Text)
        Promedio22.Text = Pusing("###,###.##", Promedio22.Text)
        Promedio33.Text = Pusing("###,###.##", Promedio33.Text)

        Promedio.Text = Pusing("###,###.##", Promedio.Text)
        
        rstEvaluaI.Close
        
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Proveedor"
    ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
    End If
    
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Camion"
    ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion1.Text + "'"
    spCamion = ZSql
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
        Dominio1.Caption = rstCamion!Patente
        rstCamion.Close
            Else
        Dominio1.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Camion"
    ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion2.Text + "'"
    spCamion = ZSql
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
        Dominio2.Caption = rstCamion!Patente
        rstCamion.Close
            Else
        Dominio2.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Camion"
    ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion3.Text + "'"
    spCamion = ZSql
    Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCamion.RecordCount > 0 Then
        Dominio3.Caption = rstCamion!Patente
        rstCamion.Close
            Else
        Dominio3.Caption = ""
    End If
    
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Chofer"
    ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer1.Text + "'"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        DesChofer1.Caption = rstChofer!Descripcion
        rstChofer.Close
            Else
        DesChofer1.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Chofer"
    ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer2.Text + "'"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        DesChofer2.Caption = rstChofer!Descripcion
        rstChofer.Close
            Else
        DesChofer2.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Chofer"
    ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer3.Text + "'"
    spChofer = ZSql
    Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
    If rstChofer.RecordCount > 0 Then
        DesChofer3.Caption = rstChofer!Descripcion
        rstChofer.Close
            Else
        DesChofer3.Caption = ""
    End If
    
End Sub

Private Sub Baja_Click()
    T$ = "Inhabilitacion de Proveedores"
    m$ = "Desea Inhabilitar al proveedor "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        EmpresaReal = WEmpresa
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0003"
        txtOdbc = "Empresa03"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0006"
        txtOdbc = "Empresa06"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0008"
        txtOdbc = "Empresa08"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = "0009"
        txtOdbc = "Empresa09"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Estado = " + "'" + "2" + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
        WEmpresa = EmpresaReal
        txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        m$ = "El Proveedor ha sido inhabilitado"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        
        Cartel.Visible = True
        
        
    End If
End Sub




Private Sub cmdAdd_Click()
    If Val(Proveedor.Text) <> 0 And Val(Mes.Text) <> 0 And Val(Ano.Text) <> 0 Then

        ZProveedor = Proveedor.Text
        ZMes = Mes.Text
        ZAno = Ano.Text
        
        Call Ceros(ZProveedor, 11)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        ZPeriodo = ZAno + ZMes
        ZClave = ZProveedor + ZMes + ZAno
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EvaluaI"
        ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
        ZSql = ZSql + " and EvaluaI.Tipo = 1"
        
        spEvaluaI = ZSql
        Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
        If rstEvaluaI.RecordCount > 0 Then
        
            rstEvaluaI.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE EvaluaI SET "
            ZSql = ZSql + " Clave = " + "'" + ZClave + "',"
            ZSql = ZSql + " Tipo = " + "'" + "1" + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Mes = " + "'" + Mes.Text + "',"
            ZSql = ZSql + " Ano = " + "'" + Ano.Text + "',"
            ZSql = ZSql + " Periodo = " + "'" + ZPeriodo + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " Vencimiento = " + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + " Evaluador = " + "'" + Evaluador.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Camion1 = " + "'" + Camion1.Text + "',"
            ZSql = ZSql + " Camion2 = " + "'" + Camion2.Text + "',"
            ZSql = ZSql + " Camion3 = " + "'" + Camion3.Text + "',"
            ZSql = ZSql + " Chofer1 = " + "'" + Chofer1.Text + "',"
            ZSql = ZSql + " Chofer2 = " + "'" + Chofer2.Text + "',"
            ZSql = ZSql + " Chofer3 = " + "'" + Chofer3.Text + "',"
            ZSql = ZSql + " DesChofer1 = " + "'" + DesChofer1.Caption + "',"
            ZSql = ZSql + " DesChofer2 = " + "'" + DesChofer2.Caption + "',"
            ZSql = ZSql + " DesChofer3 = " + "'" + DesChofer3.Caption + "',"
            ZSql = ZSql + " Dominio1 = " + "'" + Dominio1.Caption + "',"
            ZSql = ZSql + " Dominio2 = " + "'" + Dominio2.Caption + "',"
            ZSql = ZSql + " Dominio3 = " + "'" + Dominio3.Caption + "',"
            ZSql = ZSql + " Punto11 = " + "'" + Punto11.Text + "',"
            ZSql = ZSql + " Punto12 = " + "'" + Punto12.Text + "',"
            ZSql = ZSql + " Punto13 = " + "'" + Punto13.Text + "',"
            ZSql = ZSql + " Punto14 = " + "'" + Punto14.Text + "',"
            ZSql = ZSql + " Punto15 = " + "'" + Punto15.Text + "',"
            ZSql = ZSql + " Punto21 = " + "'" + Punto21.Text + "',"
            ZSql = ZSql + " Punto22 = " + "'" + Punto22.Text + "',"
            ZSql = ZSql + " Punto23 = " + "'" + Punto23.Text + "',"
            ZSql = ZSql + " Punto24 = " + "'" + Punto24.Text + "',"
            ZSql = ZSql + " Punto25 = " + "'" + Punto25.Text + "',"
            ZSql = ZSql + " Punto31 = " + "'" + Punto31.Text + "',"
            ZSql = ZSql + " Punto32 = " + "'" + Punto32.Text + "',"
            ZSql = ZSql + " Punto33 = " + "'" + Punto33.Text + "',"
            ZSql = ZSql + " Punto34 = " + "'" + Punto34.Text + "',"
            ZSql = ZSql + " Punto35 = " + "'" + Punto35.Text + "',"
            ZSql = ZSql + " Promedio1 = " + "'" + Promedio1.Text + "',"
            ZSql = ZSql + " Promedio2 = " + "'" + Promedio2.Text + "',"
            ZSql = ZSql + " Promedio3 = " + "'" + Promedio3.Text + "',"
            ZSql = ZSql + " Promedio4 = " + "'" + Promedio4.Text + "',"
            ZSql = ZSql + " Promedio5 = " + "'" + Promedio5.Text + "',"
            ZSql = ZSql + " Promedio11 = " + "'" + Promedio11.Text + "',"
            ZSql = ZSql + " Promedio22 = " + "'" + Promedio22.Text + "',"
            ZSql = ZSql + " Promedio33 = " + "'" + Promedio33.Text + "',"
            ZSql = ZSql + " Promedio = " + "'" + Promedio.Text + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO EvaluaI ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Mes ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Vencimiento ,"
            ZSql = ZSql + "Evaluador ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Camion1 ,"
            ZSql = ZSql + "Camion2 ,"
            ZSql = ZSql + "Camion3 ,"
            ZSql = ZSql + "Chofer1 ,"
            ZSql = ZSql + "Chofer2 ,"
            ZSql = ZSql + "Chofer3 ,"
            ZSql = ZSql + "DesChofer1 ,"
            ZSql = ZSql + "DesChofer2 ,"
            ZSql = ZSql + "DesChofer3 ,"
            ZSql = ZSql + "Dominio1 ,"
            ZSql = ZSql + "Dominio2 ,"
            ZSql = ZSql + "Dominio3 ,"
            ZSql = ZSql + "Punto11 ,"
            ZSql = ZSql + "Punto12 ,"
            ZSql = ZSql + "Punto13 ,"
            ZSql = ZSql + "Punto14 ,"
            ZSql = ZSql + "Punto15 ,"
            ZSql = ZSql + "Punto21 ,"
            ZSql = ZSql + "Punto22 ,"
            ZSql = ZSql + "Punto23 ,"
            ZSql = ZSql + "Punto24 ,"
            ZSql = ZSql + "Punto25 ,"
            ZSql = ZSql + "Punto31 ,"
            ZSql = ZSql + "Punto32 ,"
            ZSql = ZSql + "Punto33 ,"
            ZSql = ZSql + "Punto34 ,"
            ZSql = ZSql + "Punto35 ,"
            ZSql = ZSql + "Promedio1 ,"
            ZSql = ZSql + "Promedio2 ,"
            ZSql = ZSql + "Promedio3 ,"
            ZSql = ZSql + "Promedio4 ,"
            ZSql = ZSql + "Promedio5 ,"
            ZSql = ZSql + "Promedio11 ,"
            ZSql = ZSql + "Promedio22 ,"
            ZSql = ZSql + "Promedio33 ,"
            ZSql = ZSql + "Promedio )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + "1" + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Mes.Text + "',"
            ZSql = ZSql + "'" + Ano.Text + "',"
            ZSql = ZSql + "'" + ZPeriodo + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + "'" + Evaluador.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Camion1.Text + "',"
            ZSql = ZSql + "'" + Camion2.Text + "',"
            ZSql = ZSql + "'" + Camion3.Text + "',"
            ZSql = ZSql + "'" + Chofer1.Text + "',"
            ZSql = ZSql + "'" + Chofer2.Text + "',"
            ZSql = ZSql + "'" + Chofer3.Text + "',"
            ZSql = ZSql + "'" + DesChofer1.Caption + "',"
            ZSql = ZSql + "'" + DesChofer2.Caption + "',"
            ZSql = ZSql + "'" + DesChofer3.Caption + "',"
            ZSql = ZSql + "'" + Dominio1.Caption + "',"
            ZSql = ZSql + "'" + Dominio2.Caption + "',"
            ZSql = ZSql + "'" + Dominio3.Caption + "',"
            ZSql = ZSql + "'" + Punto11.Text + "',"
            ZSql = ZSql + "'" + Punto12.Text + "',"
            ZSql = ZSql + "'" + Punto13.Text + "',"
            ZSql = ZSql + "'" + Punto14.Text + "',"
            ZSql = ZSql + "'" + Punto15.Text + "',"
            ZSql = ZSql + "'" + Punto21.Text + "',"
            ZSql = ZSql + "'" + Punto22.Text + "',"
            ZSql = ZSql + "'" + Punto23.Text + "',"
            ZSql = ZSql + "'" + Punto24.Text + "',"
            ZSql = ZSql + "'" + Punto25.Text + "',"
            ZSql = ZSql + "'" + Punto31.Text + "',"
            ZSql = ZSql + "'" + Punto32.Text + "',"
            ZSql = ZSql + "'" + Punto33.Text + "',"
            ZSql = ZSql + "'" + Punto34.Text + "',"
            ZSql = ZSql + "'" + Punto35.Text + "',"
            ZSql = ZSql + "'" + Promedio1.Text + "',"
            ZSql = ZSql + "'" + Promedio2.Text + "',"
            ZSql = ZSql + "'" + Promedio3.Text + "',"
            ZSql = ZSql + "'" + Promedio4.Text + "',"
            ZSql = ZSql + "'" + Promedio5.Text + "',"
            ZSql = ZSql + "'" + Promedio11.Text + "',"
            ZSql = ZSql + "'" + Promedio22.Text + "',"
            ZSql = ZSql + "'" + Promedio33.Text + "',"
            ZSql = ZSql + "'" + Promedio.Text + "')"
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        Call CmdLimpiar_Click
        Proveedor.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Val(Proveedor.Text) <> 0 And Val(Mes.Text) <> 0 And Val(Ano.Text) <> 0 Then

        ZProveedor = Proveedor.Text
        ZMes = Mes.Text
        ZAno = Ano.Text
        
        Call Ceros(ZProveedor, 11)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        ZClave = ZProveedor + ZMes + ZAno
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EvaluaI"
        ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
        ZSql = ZSql + " and EvaluaI.Tipo = 1"
        
        spEvaluaI = ZSql
        Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
        If rstEvaluaI.RecordCount > 0 Then
            rstEvaluaI.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE EvaluaI"
                ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                ZSql = ZSql + " and EvaluaI.Tipo = 1"
                spEvaluaI = ZSql
                Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    
    Proveedor.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Mes.Text = ""
    Ano.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Evaluador.Text = ""
    Observaciones.Text = ""
    Chofer1.Text = ""
    Chofer2.Text = ""
    Chofer3.Text = ""
    Camion1.Text = ""
    Camion2.Text = ""
    Camion3.Text = ""
    Punto11.Text = ""
    Punto12.Text = ""
    Punto13.Text = ""
    Punto14.Text = ""
    Punto15.Text = ""
    Punto21.Text = ""
    Punto22.Text = ""
    Punto23.Text = ""
    Punto24.Text = ""
    Punto25.Text = ""
    Punto31.Text = ""
    Punto32.Text = ""
    Punto33.Text = ""
    Punto34.Text = ""
    Punto35.Text = ""
    
    DesChofer1.Caption = ""
    DesChofer2.Caption = ""
    DesChofer3.Caption = ""
    
    Dominio1.Caption = ""
    Dominio2.Caption = ""
    Dominio3.Caption = ""
    
    Promedio1.Text = ""
    Promedio2.Text = ""
    Promedio3.Text = ""
    Promedio4.Text = ""
    Promedio5.Text = ""
    
    Promedio11.Text = ""
    Promedio22.Text = ""
    Promedio33.Text = ""
    
    Promedio.Text = ""
    
    Cartel.Visible = False
    
    Proveedor.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgEvaluaTransporte.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()

    ZProveedor = Proveedor.Text
    ZMes = Mes.Text
    ZAno = Ano.Text
    
    Call Ceros(ZProveedor, 11)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZClave = ZProveedor + ZMes + ZAno

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave < " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 1"
    ZSql = ZSql + " Order by EvaluaI.Clave"
    
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
        With rstEvaluaI
            .MoveLast
            Proveedor.Text = rstEvaluaI!Proveedor
            Ano.Text = rstEvaluaI!Ano
            Mes.Text = rstEvaluaI!Mes
        End With
        rstEvaluaI.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
    End If
    
End Sub

Private Sub Evaluador_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Evaluador.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Vencimiento.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Punto11.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
End Sub

Private Sub Punto11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto12.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto11.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto13.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto12.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto14.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto13.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto15.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto14.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto21.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto15.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Punto21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto22.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto21.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto23.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto22.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto24.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto23.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto25.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto24.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto31.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto25.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub






Private Sub Punto31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto32.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto31.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto33.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto32.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto33_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto34.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto33.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto34_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto35.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto34.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto35_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Promedio
        Punto11.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto35.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Calcula_Promedio()

    ZSuma1 = Val(Punto11.Text) + Val(Punto21.Text) + Val(Punto31.Text)
    ZSuma2 = Val(Punto12.Text) + Val(Punto22.Text) + Val(Punto32.Text)
    ZSuma3 = Val(Punto13.Text) + Val(Punto23.Text) + Val(Punto33.Text)
    ZSuma4 = Val(Punto14.Text) + Val(Punto24.Text) + Val(Punto34.Text)
    ZSuma5 = Val(Punto15.Text) + Val(Punto25.Text) + Val(Punto35.Text)
    
    ZCanti = 0
    If Val(Camion1.Text) <> 0 Then
        ZCanti = ZCanti + 1
    End If
    If Val(Camion2.Text) <> 0 Then
        ZCanti = ZCanti + 1
    End If
    If Val(Camion3.Text) <> 0 Then
        ZCanti = ZCanti + 1
    End If
    
    If ZCanti <> 0 Then
        ZPromedio1 = ZSuma1 / ZCanti
        ZPromedio2 = ZSuma2 / ZCanti
        ZPromedio3 = ZSuma3 / ZCanti
        ZPromedio4 = ZSuma4 / ZCanti
        ZPromedio5 = ZSuma5 / ZCanti
            Else
        ZPromedio1 = 0
        ZPromedio2 = 0
        ZPromedio3 = 0
        ZPromedio4 = 0
        ZPromedio5 = 0
    End If
    
    Call Redondeo(ZPromedio1)
    Call Redondeo(ZPromedio2)
    Call Redondeo(ZPromedio3)
    Call Redondeo(ZPromedio4)
    Call Redondeo(ZPromedio5)
    
    Promedio1.Text = Str$(ZPromedio1)
    Promedio2.Text = Str$(ZPromedio2)
    Promedio3.Text = Str$(ZPromedio3)
    Promedio4.Text = Str$(ZPromedio4)
    Promedio5.Text = Str$(ZPromedio5)
    
    ZSuma1 = (Val(Punto11.Text) * 0.1) + (Val(Punto12.Text) * 0.2) + (Val(Punto13.Text) * 0.2) + (Val(Punto14.Text) * 0.2) + (Val(Punto15.Text) * 0.3)
    ZSuma2 = (Val(Punto21.Text) * 0.1) + (Val(Punto22.Text) * 0.2) + (Val(Punto23.Text) * 0.2) + (Val(Punto24.Text) * 0.2) + (Val(Punto25.Text) * 0.3)
    ZSuma3 = (Val(Punto31.Text) * 0.1) + (Val(Punto32.Text) * 0.2) + (Val(Punto23.Text) * 0.2) + (Val(Punto34.Text) * 0.2) + (Val(Punto35.Text) * 0.3)
    
    ZPromedio11 = ZSuma1
    ZPromedio22 = ZSuma2
    ZPromedio33 = ZSuma3
    
    Call Redondeo(ZPromedio11)
    Call Redondeo(ZPromedio22)
    Call Redondeo(ZPromedio33)
    
    Promedio11.Text = Str$(ZPromedio11)
    Promedio22.Text = Str$(ZPromedio22)
    Promedio33.Text = Str$(ZPromedio33)
    
    ZPromedio = (Val(Promedio1.Text) * 0.1) + (Val(Promedio2.Text) * 0.2) + (Val(Promedio3.Text) * 0.2) + (Val(Promedio4.Text) * 0.2) + (Val(Promedio5.Text) * 0.3)
    Promedio.Text = Str$(ZPromedio)
    
    If Val(Punto11.Text) <> 0 Then
        Punto11.Text = Pusing("###,###.##", Punto11.Text)
            Else
        Punto11.Text = ""
    End If
    If Val(Punto12.Text) <> 0 Then
        Punto12.Text = Pusing("###,###.##", Punto12.Text)
            Else
        Punto12.Text = ""
    End If
    If Val(Punto13.Text) <> 0 Then
        Punto13.Text = Pusing("###,###.##", Punto13.Text)
            Else
        Punto13.Text = ""
    End If
    If Val(Punto14.Text) <> 0 Then
        Punto14.Text = Pusing("###,###.##", Punto14.Text)
            Else
        Punto14.Text = ""
    End If
    If Val(Punto15.Text) <> 0 Then
        Punto15.Text = Pusing("###,###.##", Punto15.Text)
            Else
        Punto15.Text = ""
    End If
    
    If Val(Punto21.Text) <> 0 Then
        Punto21.Text = Pusing("###,###.##", Punto21.Text)
            Else
        Punto21.Text = ""
    End If
    If Val(Punto22.Text) <> 0 Then
        Punto22.Text = Pusing("###,###.##", Punto22.Text)
            Else
        Punto22.Text = ""
    End If
    If Val(Punto23.Text) <> 0 Then
        Punto23.Text = Pusing("###,###.##", Punto23.Text)
            Else
        Punto23.Text = ""
    End If
    If Val(Punto24.Text) <> 0 Then
        Punto24.Text = Pusing("###,###.##", Punto24.Text)
            Else
        Punto24.Text = ""
    End If
    If Val(Punto25.Text) <> 0 Then
        Punto25.Text = Pusing("###,###.##", Punto25.Text)
            Else
        Punto25.Text = ""
    End If
    
    If Val(Punto31.Text) <> 0 Then
        Punto31.Text = Pusing("###,###.##", Punto31.Text)
            Else
        Punto31.Text = ""
    End If
    If Val(Punto32.Text) <> 0 Then
        Punto32.Text = Pusing("###,###.##", Punto32.Text)
            Else
        Punto32.Text = ""
    End If
    If Val(Punto33.Text) <> 0 Then
        Punto33.Text = Pusing("###,###.##", Punto33.Text)
            Else
        Punto33.Text = ""
    End If
    If Val(Punto34.Text) <> 0 Then
        Punto34.Text = Pusing("###,###.##", Punto34.Text)
            Else
        Punto34.Text = ""
    End If
    If Val(Punto35.Text) <> 0 Then
        Punto35.Text = Pusing("###,###.##", Punto35.Text)
            Else
        Punto35.Text = ""
    End If
    
    Promedio1.Text = Pusing("###,###.##", Promedio1.Text)
    Promedio2.Text = Pusing("###,###.##", Promedio2.Text)
    Promedio3.Text = Pusing("###,###.##", Promedio3.Text)
    Promedio4.Text = Pusing("###,###.##", Promedio4.Text)
    Promedio5.Text = Pusing("###,###.##", Promedio5.Text)
    
    Promedio11.Text = Pusing("###,###.##", Promedio11.Text)
    Promedio22.Text = Pusing("###,###.##", Promedio22.Text)
    Promedio33.Text = Pusing("###,###.##", Promedio33.Text)
    
    Promedio.Text = Pusing("###,###.##", Promedio.Text)

End Sub












Private Sub Proveedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM Proveedor"
        ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = rstProveedor!Nombre
            WEstado = IIf(IsNull(rstProveedor!Estado), "0", rstProveedor!Estado)
            rstProveedor.Close
            
            If WEstado = 2 Then
                Cartel.Visible = True
                    Else
                Cartel.Visible = False
            End If
            
            
        
        
        
            ZLugar = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Camion"
            ZSql = ZSql + " Where Camion.Proveedor = " + "'" + Proveedor.Text + "'"
            ZSql = ZSql + " Order by Camion.Codigo"
            spCamion = ZSql
            Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
            If rstCamion.RecordCount > 0 Then
                With rstCamion
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZLugar = ZLugar + 1
                            Select Case ZLugar
                                Case 1
                                    Camion1.Text = rstCamion!Codigo
                                    Chofer1.Text = rstCamion!Chofer
                                Case 2
                                    Camion2.Text = rstCamion!Codigo
                                    Chofer2.Text = rstCamion!Chofer
                                Case 3
                                    Camion3.Text = rstCamion!Codigo
                                    Chofer3.Text = rstCamion!Chofer
                                Case Else
                            End Select
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCamion.Close
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Camion"
            ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion1.Text + "'"
            spCamion = ZSql
            Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
            If rstCamion.RecordCount > 0 Then
                Dominio1.Caption = rstCamion!Patente
                rstCamion.Close
                    Else
                Dominio1.Caption = ""
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Camion"
            ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion2.Text + "'"
            spCamion = ZSql
            Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
            If rstCamion.RecordCount > 0 Then
                Dominio2.Caption = rstCamion!Patente
                rstCamion.Close
                    Else
                Dominio2.Caption = ""
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Camion"
            ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion3.Text + "'"
            spCamion = ZSql
            Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
            If rstCamion.RecordCount > 0 Then
                Dominio3.Caption = rstCamion!Patente
                rstCamion.Close
                    Else
                Dominio3.Caption = ""
            End If
            
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Chofer"
            ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer1.Text + "'"
            spChofer = ZSql
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                DesChofer1.Caption = rstChofer!Descripcion
                rstChofer.Close
                    Else
                DesChofer1.Caption = ""
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Chofer"
            ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer2.Text + "'"
            spChofer = ZSql
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                DesChofer2.Caption = rstChofer!Descripcion
                rstChofer.Close
                    Else
                DesChofer2.Caption = ""
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Chofer"
            ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer3.Text + "'"
            spChofer = ZSql
            Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
            If rstChofer.RecordCount > 0 Then
                DesChofer3.Caption = rstChofer!Descripcion
                rstChofer.Close
                    Else
                DesChofer3.Caption = ""
            End If
            
            ZCanti = 0
            If Val(Camion1.Text) <> 0 Then
                ZCanti = ZCanti + 1
            End If
            If Val(Camion2.Text) <> 0 Then
                ZCanti = ZCanti + 1
            End If
            If Val(Camion3.Text) <> 0 Then
                ZCanti = ZCanti + 1
            End If
            
            
            If ZCanti = 0 Then
                m$ = "No existe camiones y choferes asociados a este proveedor"
                A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
            End If
            
            
            
            
            
            
            
            
            
            
            Mes.SetFocus
                Else
            DesProveedor.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Mes.Text) >= 1 And Val(Mes.Text) <= 12 Then
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Mes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Ano.Text) >= 1990 And Val(Mes.Text) <= 2100 Then

            ZProveedor = Proveedor.Text
            ZMes = Mes.Text
            ZAno = Ano.Text
            
            Call Ceros(ZProveedor, 11)
            Call Ceros(ZMes, 2)
            Call Ceros(ZAno, 4)
            
            ZClave = ZProveedor + ZMes + ZAno
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM EvaluaI"
            ZSql = ZSql + " Where EvaluaI.Clave = " + "'" + ZClave + "'"
            ZSql = ZSql + " and EvaluaI.Tipo = 1"
            
            spEvaluaI = ZSql
            Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
            If rstEvaluaI.RecordCount > 0 Then
            
                rstEvaluaI.Close
                Call Imprime_Datos
                
                    Else
                    
                ZDesProveedor = DesProveedor.Caption
                ZProveedor = Proveedor.Text
                ZMes = Mes.Text
                ZAno = Ano.Text
                CmdLimpiar_Click
                Proveedor.Text = ZProveedor
                Mes.Text = ZMes
                Ano.Text = ZAno
                DesProveedor.Caption = ZDesProveedor
                
                ZLugar = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Camion"
                ZSql = ZSql + " Where Camion.Proveedor = " + "'" + Proveedor.Text + "'"
                ZSql = ZSql + " Order by Camion.Codigo"
                spCamion = ZSql
                Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
                If rstCamion.RecordCount > 0 Then
                    With rstCamion
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                ZLugar = ZLugar + 1
                                Select Case ZLugar
                                    Case 1
                                        Camion1.Text = rstCamion!Codigo
                                        Chofer1.Text = rstCamion!Chofer
                                    Case 2
                                        Camion2.Text = rstCamion!Codigo
                                        Chofer2.Text = rstCamion!Chofer
                                    Case 3
                                        Camion3.Text = rstCamion!Codigo
                                        Chofer3.Text = rstCamion!Chofer
                                    Case Else
                                End Select
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCamion.Close
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Camion"
                ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion1.Text + "'"
                spCamion = ZSql
                Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
                If rstCamion.RecordCount > 0 Then
                    Dominio1.Caption = rstCamion!Patente
                    rstCamion.Close
                        Else
                    Dominio1.Caption = ""
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Camion"
                ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion2.Text + "'"
                spCamion = ZSql
                Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
                If rstCamion.RecordCount > 0 Then
                    Dominio2.Caption = rstCamion!Patente
                    rstCamion.Close
                        Else
                    Dominio2.Caption = ""
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Camion"
                ZSql = ZSql & " Where Camion.Codigo = " + "'" + Camion3.Text + "'"
                spCamion = ZSql
                Set rstCamion = db.OpenRecordset(spCamion, dbOpenSnapshot, dbSQLPassThrough)
                If rstCamion.RecordCount > 0 Then
                    Dominio3.Caption = rstCamion!Patente
                    rstCamion.Close
                        Else
                    Dominio3.Caption = ""
                End If
                
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Chofer"
                ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer1.Text + "'"
                spChofer = ZSql
                Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
                If rstChofer.RecordCount > 0 Then
                    DesChofer1.Caption = rstChofer!Descripcion
                    rstChofer.Close
                        Else
                    DesChofer1.Caption = ""
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Chofer"
                ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer2.Text + "'"
                spChofer = ZSql
                Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
                If rstChofer.RecordCount > 0 Then
                    DesChofer2.Caption = rstChofer!Descripcion
                    rstChofer.Close
                        Else
                    DesChofer2.Caption = ""
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Chofer"
                ZSql = ZSql & " Where Chofer.Codigo = " + "'" + Chofer3.Text + "'"
                spChofer = ZSql
                Set rstChofer = db.OpenRecordset(spChofer, dbOpenSnapshot, dbSQLPassThrough)
                If rstChofer.RecordCount > 0 Then
                    DesChofer3.Caption = rstChofer!Descripcion
                    rstChofer.Close
                        Else
                    DesChofer3.Caption = ""
                End If
                
                ZCanti = 0
                If Val(Camion1.Text) <> 0 Then
                    ZCanti = ZCanti + 1
                End If
                If Val(Camion2.Text) <> 0 Then
                    ZCanti = ZCanti + 1
                End If
                If Val(Camion3.Text) <> 0 Then
                    ZCanti = ZCanti + 1
                End If
                
                ZSql = ""
                ZSql = ZSql & "Select *"
                ZSql = ZSql & " FROM Proveedor"
                ZSql = ZSql & " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    DesProveedor.Caption = rstProveedor!Nombre
                    WEstado = IIf(IsNull(rstProveedor!Estado), "0", rstProveedor!Estado)
                    If WEstado = 2 Then
                        Cartel.Visible = True
                            Else
                        Cartel.Visible = False
                    End If
                    rstProveedor.Close
                End If
                
                If ZCanti = 0 Then
                    m$ = "No existe camiones y choferes asociados a este proveedor"
                    A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
                End If
                
                
            End If
            
            Evaluador.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Proveedores"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor.Proveedor"
            spProveedor = Sql1 + Sql2 + Sql3
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstProveedor!Proveedor
                            Pantalla.Col = 2
                            Pantalla.Text = rstProveedor!Nombre
                            IngresaItem = rstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_Keypress(13)
            
        Case Else
    End Select
    
End Sub


Private Sub Siguiente_Click()
    ZProveedor = Proveedor.Text
    ZMes = Mes.Text
    ZAno = Ano.Text
    
    Call Ceros(ZProveedor, 11)
    Call Ceros(ZMes, 2)
    Call Ceros(ZAno, 4)
    
    ZClave = ZProveedor + ZMes + ZAno

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EvaluaI"
    ZSql = ZSql + " Where EvaluaI.Clave > " + "'" + ZClave + "'"
    ZSql = ZSql + " and EvaluaI.Tipo = 1"
    ZSql = ZSql + " Order by EvaluaI.Clave"
            
    spEvaluaI = ZSql
    Set rstEvaluaI = db.OpenRecordset(spEvaluaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstEvaluaI.RecordCount > 0 Then
        With rstEvaluaI
            .MoveFirst
            Proveedor.Text = rstEvaluaI!Proveedor
            Ano.Text = rstEvaluaI!Ano
            Mes.Text = rstEvaluaI!Mes
        End With
        rstEvaluaI.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Evaluaciones")
    End If

End Sub

Sub Form_Load()

    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Mes.Text = ""
    Ano.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Evaluador.Text = ""
    Observaciones.Text = ""
    Chofer1.Text = ""
    Chofer2.Text = ""
    Chofer3.Text = ""
    Camion1.Text = ""
    Camion2.Text = ""
    Camion3.Text = ""
    Punto11.Text = ""
    Punto12.Text = ""
    Punto13.Text = ""
    Punto14.Text = ""
    Punto15.Text = ""
    Punto21.Text = ""
    Punto22.Text = ""
    Punto23.Text = ""
    Punto24.Text = ""
    Punto25.Text = ""
    Punto31.Text = ""
    Punto32.Text = ""
    Punto33.Text = ""
    Punto34.Text = ""
    Punto35.Text = ""
    
    DesChofer1.Caption = ""
    DesChofer2.Caption = ""
    DesChofer3.Caption = ""
    
    Dominio1.Caption = ""
    Dominio2.Caption = ""
    Dominio3.Caption = ""
    
    Promedio1.Text = ""
    Promedio2.Text = ""
    Promedio3.Text = ""
    Promedio4.Text = ""
    Promedio5.Text = ""
    
    Promedio.Text = ""
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM Proveedor"
            Sql3 = " Order by Proveedor.Proveedor"
            spProveedor = Sql1 + Sql2 + Sql3
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            DA = Len(rstProveedor!Nombre) - WEspacios
                            For aa = 1 To DA + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstProveedor!Nombre, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstProveedor!Proveedor
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstProveedor!Nombre
                                    IngresaItem = rstProveedor!Proveedor
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
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedor"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
    End Select
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    Select Case XIndice
        Case 0
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Codigo"
                        Pantalla.ColWidth(Ciclo) = 1500
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 6000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = Pantalla.Text
        WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        WTitulo(Ciclo).Width = Pantalla.CellWidth
        WTitulo(Ciclo).Height = Pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub





