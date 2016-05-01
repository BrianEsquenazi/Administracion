VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHoja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "6"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9240
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Index           =   8
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1440
      TabIndex        =   78
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame PantaObserva 
      BackColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   480
      TabIndex        =   74
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton CierraObserva 
         Caption         =   "Cierra"
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
         Left            =   3960
         TabIndex        =   77
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox ImpreObserva 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   76
         Top             =   1110
         Width           =   8175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   960
         TabIndex        =   75
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame PantaTeorico 
      Height          =   1335
      Left            =   1560
      TabIndex        =   68
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox FabricaIII 
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
         HideSelection   =   0   'False
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   73
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox FabricaII 
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   71
         Text            =   " "
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Fabrica 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   69
         Text            =   " "
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Lotes Autorizados"
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
         TabIndex        =   72
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Rendimiento teorico"
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
         TabIndex        =   70
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   615
      Left            =   10560
      TabIndex        =   43
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin RichTextLib.RichTextBox Agenda 
      Height          =   615
      Left            =   9480
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   327680
      ScrollBars      =   3
      RightMargin     =   8900
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"hoja.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   4815
      Left            =   120
      TabIndex        =   67
      Top             =   1560
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.TextBox NroPedido 
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
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   65
      Text            =   " "
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer 
      Left            =   11160
      Top             =   2040
   End
   Begin VB.TextBox Salida 
      Height          =   285
      Left            =   11520
      TabIndex        =   64
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox HoraInicio 
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
      MaxLength       =   10
      TabIndex        =   59
      Text            =   " "
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox HoraFinal 
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
      Left            =   10680
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   58
      Text            =   " "
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Reserva 
      Caption         =   "Reserva Hoja"
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
      TabIndex        =   57
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox VersionIII 
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
      Left            =   11160
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   56
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox VersionII 
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
      Left            =   9120
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   55
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox VersionI 
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
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   54
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Equipo 
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
      Left            =   10680
      MaxLength       =   2
      TabIndex        =   49
      Text            =   " "
      Top             =   840
      Width           =   615
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
      Index           =   5
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   6960
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
      Index           =   4
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   6840
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   6840
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
      Index           =   2
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   6840
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
      Index           =   3
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   6960
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
      Index           =   6
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   7200
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
      Index           =   7
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partidas"
      Height          =   1815
      Left            =   6480
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   2655
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   38
         Top             =   720
         Width           =   975
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   37
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox WLote3 
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
         MaxLength       =   6
         TabIndex        =   36
         Top             =   1440
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
         Left            =   1200
         TabIndex        =   35
         Top             =   720
         Width           =   855
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
         Left            =   1200
         TabIndex        =   34
         Top             =   1080
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
         Left            =   1200
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox WControl1 
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
         TabIndex        =   32
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox WControl2 
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
         TabIndex        =   31
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox WControl3 
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
         TabIndex        =   30
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label dada 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
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
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
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
         Left            =   1200
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton Block1 
      Caption         =   "Ver Block de Notas"
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
      TabIndex        =   27
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton Block2 
      Caption         =   "Cerrar Block"
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
      TabIndex        =   26
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox observaciones 
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
      MaxLength       =   50
      TabIndex        =   25
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox Pedido 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6240
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra1 
      Height          =   1935
      Left            =   9240
      TabIndex        =   22
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3413
      _Version        =   327680
      Rows            =   150
      Cols            =   4
   End
   Begin VB.TextBox Stock 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   2415
      Left            =   9240
      TabIndex        =   20
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      _Version        =   327680
      Rows            =   100
      Cols            =   4
   End
   Begin VB.CommandButton Reimpresion 
      Caption         =   "ReImpre."
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
      TabIndex        =   19
      Top             =   7680
      Width           =   975
   End
   Begin MSMask.MaskEdBox fechaIng 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
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
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
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
   Begin VB.TextBox Real 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   4
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Teorico 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9600
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ImpreHoja.rpt"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   7080
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
      Height          =   1500
      Left            =   3360
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox Hoja 
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
      Left            =   2040
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
      Left            =   120
      TabIndex        =   9
      Top             =   6480
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
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7080
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      Top             =   6600
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
      ItemData        =   "hoja.frx":007C
      Left            =   3360
      List            =   "hoja.frx":0083
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   8415
   End
   Begin MSMask.MaskEdBox FechaInicio 
      Height          =   285
      Left            =   3360
      TabIndex        =   60
      Top             =   1200
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
   Begin MSMask.MaskEdBox FechaFinal 
      Height          =   285
      Left            =   9240
      TabIndex        =   61
      Top             =   1200
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
   Begin VB.Label LabelPedido 
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
      Left            =   3600
      TabIndex        =   66
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label23 
      Caption         =   "Fecha y Hora Inicio  Produccion"
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
      TabIndex        =   63
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label24 
      Caption         =   "Fecha y Hora de Envasamiento"
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
      Left            =   6240
      TabIndex        =   62
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label17 
      Caption         =   "Especif."
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
      Left            =   9960
      TabIndex        =   53
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Procedim. "
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
      Left            =   8160
      TabIndex        =   52
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Version Formula"
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
      TabIndex        =   51
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Equipo"
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
      Left            =   9840
      TabIndex        =   50
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   5520
      TabIndex        =   24
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha Ingreso"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Rendimiento Real"
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
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Rendimiento teorico"
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
      Width           =   1815
   End
   Begin VB.Label DesProducto 
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
      Left            =   3600
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Producto"
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
      TabIndex        =   12
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
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Hoja de Produccion"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "PrgHoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Clave As String
Private WAnterior As Integer
Private Tipo As String

Dim ZZDescriII(100) As String
Dim ZZEnsayo(100) As String
Dim ZZStd(100, 10) As String

Private Existe  As String
Private Auxi1 As String
Private Auxi2 As String
Private XIndice As Integer
Private WImpre As String
Private Cantidad As String
Private XCantidad As Double
Private Auxiliar(100, 7) As String
Private AuxiliarII(100) As String
Private AuxiliarIII(100, 7) As String
Dim vers As String
Dim fechaversio As String

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposision As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPrecio As Recordset
Dim spPrecio As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstImpreHoja As Recordset
Dim spImpreHoja As String
Dim rstEfluentes As Recordset
Dim spEfluentes As String
Dim rstImpreHojaII As Recordset
Dim spImpreHojaII As String
Dim rstImpreCarga As Recordset
Dim spImpreCarga As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstCargaIII As Recordset
Dim spCargaIII As String
Dim rstCargaV As Recordset
Dim spCargaV As String
Dim rstImpreCargaI As Recordset
Dim spImpreCargaI As String

Dim XParam As String
Dim LeeHoja As String
Dim XSaldo As Double
Dim Impre(10, 2) As Double
Dim WCosto1 As String
Dim WCosto2 As String
Dim WCosto3 As String
Private ZLote(100, 7) As String
Dim WSaldo1 As Double
Dim WSaldo2 As Double
Dim WSaldo3 As Double
Dim XSaldo1 As String
Dim XSaldo2 As String
Dim XSaldo3 As String
Dim WEstado As String
Private BajaLote(3, 2) As String
Private WControla As String
Private WSaldoant As Double
Private ZCantidad As Double
Private WExiste As String
Dim ZProceso As Integer
Dim LoteBusqueda As String
Dim ZSaldo As Double
Dim ZMetodo As String
Dim ZEfluentes As String
Dim ZVersionI As String
Dim ZVersionII As String
Dim ZVersionIII As String
Dim ZDesEfluentesI As String
Dim ZDesEfluentesII As String
Dim ZZCanti(10) As String
Dim ZZLote(10) As String
Dim ZZMetodo(10) As String
Dim ZZEspecificacion(20) As String
Dim ZZDesEnsayo(10) As String
Dim ZCompo(100, 10) As String
Dim WVersionI As Integer
Dim WVersionII As Integer
Dim WVersionIII As Integer
Dim ZFarma(1000, 5) As String
Dim ZHumedad(100) As String
Dim ZComparaI(100, 2) As String
Dim ZComparaII(100, 2) As String
Dim ZImpreCarga(200, 6) As String
Dim ZImpreCargaI(100, 20) As String
Dim ZImpreMetodo(100) As String

Dim WEscrito As Integer
Dim WTeorico As String
Dim ZZEquipo As Integer
Dim ZZDescripcionI As String
Dim ZZDescripcionII As String
Dim ZZCantidad As Integer
Dim ZDesTerminado As String
Dim ZZArticulo As String
Dim ZZDesvio As Double
Dim ZZSaldo As Double
Dim ZZCantidadHoja As Double
Dim ZZBloqueo As Double
Dim ZZTipoBloqueo As String
Dim ZZCodigoBloqueo As String

Dim XXSaldo As Double
Dim ZFechaVto As String
Dim XMes As String
Dim XAno As String
Dim ZVto As String
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim QSaldo As Double

Dim ZArti(100, 10) As String
Dim Empe(12, 10) As String
Dim WWArticulo As String
Dim WWSaldo As Double
Dim WWVencido As Double

Dim WWTipo As String
Dim WWLugar As Integer
Dim WWTerminado As String
Dim WWCantidad As String

Private Sub BAJAOK_Click()

    Sql1 = "DELETE Hoja"
    Sql2 = " Where Hoja = " + "'" + Hoja.Text + "'"
    spHoja = Sql1 + Sql2
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Block1_Click()

    On Error GoTo WError

    Agenda.LoadFile "blanco.rtf", 0
    Agenda.LoadFile "H" + Hoja.Text + ".rtf", 0
    Agenda.Visible = True
    Block1.Visible = False
    Block2.Visible = True
    Agenda.Height = 6700
    Agenda.Left = 840
    Agenda.Top = 720
    Agenda.Width = 9375
    Agenda.SetFocus
    
WError:
    Resume Next
    
End Sub

Private Sub Block2_Click()
    Agenda.SaveFile "H" + Hoja.Text + ".rtf", 0
    Agenda.Visible = False
    Block1.Visible = True
    Block2.Visible = False
End Sub

Private Sub CierraObserva_Click()
    PantaObserva.Visible = False
    Teorico.SetFocus
End Sub

Private Sub cmdClose_Click()
    With rstEtiqueta
        .Close
    End With
    PrgHoja.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

    Dim WSalva(5000, 2) As String
    
    Erase WSalva
    WLugar = 0

    WRenglon = "1"
    WDesde = "20050101"
    WHasta = "20051231"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Renglon = " + "'" + WRenglon + "'"
    ZSql = ZSql + " and Hoja.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Hoja.FechaOrd <= " + "'" + WHasta + "'"
    
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WLugar = WLugar + 1
                    
                    WSalva(WLugar, 1) = Str$(rstHoja!Hoja)
                    WSalva(WLugar, 2) = rstHoja!Producto
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To WLugar
    
        WHoja = WSalva(Ciclo, 1)
        WProducto = WSalva(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Producto = " + "'" + WProducto + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + WHoja + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Call cmdClose_Click

End Sub

Private Sub Command9_Click()
    Call Impresion_Farma
End Sub

Private Sub Command2_Click()
    Call Impresion
End Sub

Private Sub Command23_Click()


    Dim WSalva(5000, 2) As String
    
    
    Dim ZZClave As String
    Dim ZZHoja As Integer
    Dim ZZRenglon As Integer
    Dim ZZFecha As String
    Dim ZZProducto As String
    Dim ZZCantidad As Double
    Dim ZZTipo As String
    Dim ZZZZLote As Integer
    Dim ZZArticulo As String
    Dim ZZTerminado As String
    
    Dim ZZTeorico As Double
    Dim ZZReal As Double
    Dim ZZFechaing As String
    Dim ZZFechaIngOrd As String
    Dim ZZWDate As String
    Dim ZZWImporte As Double
    Dim ZZMarca As String
    Dim ZZSaldo As Double
    Dim ZZLote1 As Integer
    Dim ZZCanti1  As Double
    
    Dim ZZLote2 As Integer
    Dim ZZCanti2 As Double
    Dim ZZLote3 As Integer
    Dim ZZCanti3 As Double
    Dim ZZCosto1 As Double
    Dim ZZCosto2 As Double
    Dim ZZCosto3 As Double
    Dim ZZMarcaAnt As String
    Dim ZZSaldoAnt As Double
    Dim ZZRealAnt As Double
    
    Dim ZZFechaOrd As String
    Dim ZZEquipo As String
    Dim ZZMarcaLabora As String
    Dim ZZEstado As String
    Dim ZZVersionI As Integer
    Dim ZZVersionII As Integer
    Dim ZZVersionIII As Integer
    Dim ZZEstadoII As String
    Dim ZZImpreArticulo As String
    Dim ZZFechaInicio As String
    
    Dim ZZHoraInicio As String
    Dim ZZFechafinal As String
    Dim ZZHoraFinal As String
    Dim ZZPorceDife As Double
    Dim ZZImpresionI As String
    Dim ZZImpresionII As String
    Dim ZZMotivoDesvio As Integer
    Dim ZZObservaDesvio As String
    Dim ZZImpreReal As Double
    Dim ZZOperario As Integer
    
    Dim ZZEstadoHoja As Integer
    Dim ZZEtapa As Integer
    Dim ZZFechaInicioEtaoa As String
    Dim ZZHoraInicioEtapa As String
    Dim ZZTimerInicioEtapa As Double
    Dim ZZAlarma As String
    Dim ZZControlI As Integer
    Dim ZZControlII As Integer
    Dim ZZDesdeI As Double
    Dim ZZHastaI As Double
    
    Dim ZZTiempoI As Double
    Dim ZZTiempoII As Double
    Dim ZZAlarmaI As String
    Dim ZZAlarmaII As String
    Dim ZZAlarmaITiempo As Double
    Dim ZZAlarmaITempe As Double
    Dim ZZTiempoIII As Double
    Dim ZZTemperatura As Double
    Dim ZZTipoEtapa As Integer
    Dim ZZEnvasamiento As String
    
    Dim ZZEquipoII As Integer
    Dim ZZDesde As Double
    Dim ZZHasta As Double
    Dim ZZLista As String
    Dim ZZSuma1 As Double
    Dim ZZSuma2 As Double
    Dim ZZSuma3 As Double
    Dim ZZSuma4 As Double
    Dim ZZSuma5 As Double
    Dim ZZSuma6 As Double
    
    Dim ZZIdentificacion As String
    Dim ZZNroPedido As Double
    Dim ZZFechaVencimiento As String
    Dim ZZOrdFechaVencimiento As String
    Dim ZZRevalida As Integer
    Dim ZZFechaRevalida As String
    Dim ZZOrdFechaRevalida As String
    Dim ZZMesesRevalida As Integer
    Dim ZZMarcaVencida As String
    Dim ZZLoteColorante As String
    
    Dim ZZTipoOri As String
    Dim ZZImpreVersion As Integer
    Dim ZZImpreFechaVersion As String
    
    
    
    
    
    ZPasa = 0
    ZCorte = ""
    ZSuma = 0
    
    Erase WSalva
    WLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Order by Clave"
    
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstHoja!Clave
                        ZSuma = 0
                    End If
                    
                    If ZCorte <> rstHoja!Clave Then
                        If ZSuma > 1 Then
                            WLugar = WLugar + 1
                            WSalva(WLugar, 1) = ZCorte
                        End If
                        ZCorte = rstHoja!Clave
                        ZSuma = 0
                    End If
                    
                    ZSuma = ZSuma + 1
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    
    
Stop
    
    For Ciclo = 1 To WLugar
    
        ZZClave = WSalva(Ciclo, 1)
        ZZClave = WSalva(7, 1)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Clave = " + "'" + ZZClave + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            
            ZZClave = IIf(IsNull(rstHoja!Clave), "", rstHoja!Clave)
            ZZZZHoja = IIf(IsNull(rstHoja!Hoja), "0", rstHoja!Hoja)
            ZZRenglon = IIf(IsNull(rstHoja!Renglon), "0", rstHoja!Renglon)
            ZZFecha = IIf(IsNull(rstHoja!Fecha), "", rstHoja!Fecha)
            ZZProducto = IIf(IsNull(rstHoja!Producto), "", rstHoja!Producto)
            ZZCantidad = IIf(IsNull(rstHoja!Cantidad), "0", rstHoja!Cantidad)
            ZZTipo = IIf(IsNull(rstHoja!Tipo), "", rstHoja!Tipo)
            ZZZZLote = IIf(IsNull(rstHoja!Lote), "0", rstHoja!Lote)
            ZZArticulo = IIf(IsNull(rstHoja!Articulo), "", rstHoja!Articulo)
            ZZTerminado = IIf(IsNull(rstHoja!Terminado), "", rstHoja!Terminado)
            
            ZZTeorico = IIf(IsNull(rstHoja!Teorico), "0", rstHoja!Teorico)
            ZZReal = IIf(IsNull(rstHoja!Real), "0", rstHoja!Real)
            ZZFechaing = IIf(IsNull(rstHoja!fechaIng), "", rstHoja!fechaIng)
            ZZFechaIngOrd = IIf(IsNull(rstHoja!FechaIngOrd), "", rstHoja!FechaIngOrd)
            ZZWDate = IIf(IsNull(rstHoja!WDate), "", rstHoja!WDate)
            ZZWImporte = IIf(IsNull(rstHoja!WImporte), "0", rstHoja!WImporte)
            ZZMarca = IIf(IsNull(rstHoja!Marca), "", rstHoja!Marca)
            ZZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
            ZZZZLote1 = IIf(IsNull(rstHoja!lote1), "0", rstHoja!lote1)
            ZZCanti1 = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
            
            ZZZZLote2 = IIf(IsNull(rstHoja!lote2), "0", rstHoja!lote2)
            ZZCanti2 = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
            ZZZZLote3 = IIf(IsNull(rstHoja!lote3), "0", rstHoja!lote3)
            ZZCanti3 = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
            ZZCosto1 = IIf(IsNull(rstHoja!Costo1), "0", rstHoja!Costo1)
            ZZCosto2 = IIf(IsNull(rstHoja!Costo2), "0", rstHoja!Costo2)
            ZZCosto3 = IIf(IsNull(rstHoja!Costo3), "0", rstHoja!Costo3)
            ZZMarcaAnt = IIf(IsNull(rstHoja!MarcaAnt), "", rstHoja!MarcaAnt)
            ZZSaldoAnt = IIf(IsNull(rstHoja!Saldoant), "0", rstHoja!Saldoant)
            ZZRealAnt = IIf(IsNull(rstHoja!realant), "0", rstHoja!realant)
            
            ZZFechaOrd = IIf(IsNull(rstHoja!FechaOrd), "", rstHoja!FechaOrd)
            ZZEquipo = IIf(IsNull(rstHoja!Equipo), "", rstHoja!Equipo)
            ZZMarcaLabora = IIf(IsNull(rstHoja!MarcaLabora), "", rstHoja!MarcaLabora)
            ZZEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            ZZVersionI = IIf(IsNull(rstHoja!VersionI), "0", rstHoja!VersionI)
            ZZVersionII = IIf(IsNull(rstHoja!VersionII), "0", rstHoja!VersionII)
            ZZVersionIII = IIf(IsNull(rstHoja!VersionIII), "0", rstHoja!VersionIII)
            ZZEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
            ZZImpreArticulo = IIf(IsNull(rstHoja!ImpreArticulo), "", rstHoja!ImpreArticulo)
            ZZFechaInicio = IIf(IsNull(rstHoja!FechaInicio), "", rstHoja!FechaInicio)
            
            ZZHoraInicio = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
            ZZFechafinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
            ZZHoraFinal = IIf(IsNull(rstHoja!HoraFinal), "", rstHoja!HoraFinal)
            ZZPorceDife = IIf(IsNull(rstHoja!PorceDife), "0", rstHoja!PorceDife)
            ZZImpresionI = IIf(IsNull(rstHoja!ImpresionI), "", rstHoja!ImpresionI)
            ZZImpresionII = IIf(IsNull(rstHoja!ImpresionII), "", rstHoja!ImpresionII)
            ZZMotivoDesvio = IIf(IsNull(rstHoja!MotivoDesvio), "0", rstHoja!MotivoDesvio)
            ZZObservaDesvio = IIf(IsNull(rstHoja!ObservaDesvio), "", rstHoja!ObservaDesvio)
            ZZImpreReal = IIf(IsNull(rstHoja!ImpreReal), "0", rstHoja!ImpreReal)
            ZZOperario = IIf(IsNull(rstHoja!Operario), "0", rstHoja!Operario)
            
            ZZEstadoHoja = IIf(IsNull(rstHoja!EstadoHoja), "0", rstHoja!EstadoHoja)
            ZZEtapa = IIf(IsNull(rstHoja!Etapa), "0", rstHoja!Etapa)
            ZZFechaInicioEtaoa = IIf(IsNull(rstHoja!FechaInicioEtaoa), "", rstHoja!FechaInicioEtaoa)
            ZZHoraInicioEtapa = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
            ZZTimerInicioEtapa = IIf(IsNull(rstHoja!TimerInicioEtapa), "0", rstHoja!TimerInicioEtapa)
            ZZAlarma = IIf(IsNull(rstHoja!Alarma), "", rstHoja!Alarma)
            ZZControlI = IIf(IsNull(rstHoja!ControlI), "0", rstHoja!ControlI)
            ZZControlII = IIf(IsNull(rstHoja!ControlII), "0", rstHoja!ControlII)
            ZZDesdeI = IIf(IsNull(rstHoja!DesdeI), "0", rstHoja!DesdeI)
            ZZHastaI = IIf(IsNull(rstHoja!HastaI), "0", rstHoja!HastaI)
            
            ZZTiempoI = IIf(IsNull(rstHoja!TiempoI), "0", rstHoja!TiempoI)
            ZZTiempoII = IIf(IsNull(rstHoja!TiempoII), "0", rstHoja!TiempoII)
            ZZAlarmaI = IIf(IsNull(rstHoja!AlarmaI), "", rstHoja!AlarmaI)
            ZZAlarmaII = IIf(IsNull(rstHoja!AlarmaII), "", rstHoja!AlarmaII)
            ZZAlarmaITiempo = IIf(IsNull(rstHoja!AlarmaITiempo), "0", rstHojaAlarmaITiempo)
            ZZAlarmaITempe = IIf(IsNull(rstHoja!AlarmaITempe), "0", rstHoja!AlarmaITempe)
            ZZTiempoIII = IIf(IsNull(rstHoja!TiempoIII), "0", rstHoja!TiempoIII)
            ZZTemperatura = IIf(IsNull(rstHoja!Temperatura), "0", rstHoja!Temperatura)
            ZZTipoEtapa = IIf(IsNull(rstHoja!TipoEtapa), "0", rstHoja!TipoEtapa)
            ZZEnvasamiento = IIf(IsNull(rstHoja!Envasamiento), "", rstHoja!Envasamiento)
            
            ZZEquipoII = IIf(IsNull(rstHoja!EquipoII), "0", rstHoja!EquipoII)
            ZZDesde = IIf(IsNull(rstHoja!Desde), "0", rstHoja!Desde)
            ZZHasta = IIf(IsNull(rstHoja!Hasta), "0", rstHoja!Hasta)
            ZZLista = IIf(IsNull(rstHoja!Lista), "", rstHoja!Lista)
            ZZSuma1 = IIf(IsNull(rstHoja!Suma1), "0", rstHoja!Suma1)
            ZZSuma2 = IIf(IsNull(rstHoja!Suma2), "0", rstHoja!Suma2)
            ZZSuma3 = IIf(IsNull(rstHoja!Suma3), "0", rstHoja!Suma3)
            ZZSuma4 = IIf(IsNull(rstHoja!Suma4), "0", rstHoja!Suma4)
            ZZSuma5 = IIf(IsNull(rstHoja!Suma5), "0", rstHoja!Suma5)
            ZZSuma6 = IIf(IsNull(rstHoja!Suma6), "0", rstHoja!Suma6)
            
            ZZIdentificacion = IIf(IsNull(rstHoja!Identificacion), "", rstHoja!Identificacion)
            ZZNroPedido = IIf(IsNull(rstHoja!NroPedido), "0", rstHoja!NroPedido)
            ZZFechaVencimiento = IIf(IsNull(rstHoja!FechaVencimiento), "", rstHoja!FechaVencimiento)
            ZZOrdFechaVencimiento = IIf(IsNull(rstHoja!OrdFechaVencimiento), "", rstHoja!OrdFechaVencimiento)
            ZZRevalida = IIf(IsNull(rstHoja!Revalida), "0", rstHoja!Revalida)
            ZZFechaRevalida = IIf(IsNull(rstHoja!FechaRevalida), "", rstHoja!FechaRevalida)
            ZZOrdFechaRevalida = IIf(IsNull(rstHoja!OrdFechaRevalida), "", rstHoja!OrdFechaRevalida)
            ZZMesesRevalida = IIf(IsNull(rstHoja!MesesRevalida), "0", rstHoja!MesesRevalida)
            ZZMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
            ZZLoteColorante = IIf(IsNull(rstHoja!LoteColorante), "", rstHoja!LoteColorante)
            
            ZZTipoOri = IIf(IsNull(rstHoja!TipoOri), "", rstHoja!TipoOri)
            ZZImpreVersion = IIf(IsNull(rstHoja!ImpreVersion), "0", rstHoja!ImpreVersion)
            ZZImpreFechaVersion = IIf(IsNull(rstHoja!ImpreFechaVersion), "", rstHoja!ImpreFechaVersion)
        
            rstHoja.Close
            
        
            ZSql = "DELETE Hoja"
            ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        
            ZSql = ""
            ZSql = ZSql & "INSERT INTO Hoja ("
            ZSql = ZSql & "Clave ,"
            ZSql = ZSql & "Hoja ,"
            ZSql = ZSql & "Renglon ,"
            ZSql = ZSql & "Fecha ,"
            ZSql = ZSql & "Producto ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Tipo ,"
            ZSql = ZSql & "Lote ,"
            ZSql = ZSql & "Articulo ,"
            ZSql = ZSql & "Terminado ,"
            ZSql = ZSql & "Teorico ,"
            ZSql = ZSql & "Real ,"
            ZSql = ZSql & "Fechaing ,"
            ZSql = ZSql & "FechaIngOrd ,"
            ZSql = ZSql & "WDate ,"
            ZSql = ZSql & "WImporte ,"
            ZSql = ZSql & "Marca ,"
            ZSql = ZSql & "Saldo ,"
            ZSql = ZSql & "Lote1 ,"
            ZSql = ZSql & "Canti1 ,"
            ZSql = ZSql & "Lote2 ,"
            ZSql = ZSql & "Canti2 ,"
            ZSql = ZSql & "Lote3 ,"
            ZSql = ZSql & "Canti3 ,"
            ZSql = ZSql & "Costo1 ,"
            ZSql = ZSql & "Costo2 ,"
            ZSql = ZSql & "Costo3 ,"
            ZSql = ZSql & "MarcaAnt ,"
            ZSql = ZSql & "SaldoAnt ,"
            ZSql = ZSql & "RealAnt ,"
            ZSql = ZSql & "FechaOrd ,"
            ZSql = ZSql & "Equipo ,"
            ZSql = ZSql & "MarcaLabora ,"
            ZSql = ZSql & "Estado ,"
            ZSql = ZSql & "VersionI ,"
            ZSql = ZSql & "VersionII ,"
            ZSql = ZSql & "VersionIII ,"
            ZSql = ZSql & "EstadoII ,"
            ZSql = ZSql & "ImpreArticulo ,"
            ZSql = ZSql & "FechaInicio ,"
            ZSql = ZSql & "HoraInicio ,"
            ZSql = ZSql & "Fechafinal ,"
            ZSql = ZSql & "HoraFinal ,"
            ZSql = ZSql & "PorceDife ,"
            ZSql = ZSql & "ImpresionI ,"
            ZSql = ZSql & "ImpresionII ,"
            ZSql = ZSql & "MotivoDesvio ,"
            ZSql = ZSql & "ObservaDesvio ,"
            ZSql = ZSql & "ImpreReal ,"
            ZSql = ZSql & "Operario ,"
            ZSql = ZSql & "EstadoHoja ,"
            ZSql = ZSql & "Etapa ,"
            ZSql = ZSql & "FechaInicioEtaoa ,"
            ZSql = ZSql & "HoraInicioEtapa ,"
            ZSql = ZSql & "TimerInicioEtapa ,"
            ZSql = ZSql & "Alarma ,"
            ZSql = ZSql & "ControlI ,"
            ZSql = ZSql & "ControlII ,"
            ZSql = ZSql & "DesdeI ,"
            ZSql = ZSql & "HastaI ,"
            ZSql = ZSql & "TiempoI ,"
            ZSql = ZSql & "TiempoII ,"
            ZSql = ZSql & "AlarmaI ,"
            ZSql = ZSql & "AlarmaII ,"
            ZSql = ZSql & "AlarmaITiempo ,"
            ZSql = ZSql & "AlarmaITempe ,"
            ZSql = ZSql & "TiempoIII ,"
            ZSql = ZSql & "Temperatura ,"
            ZSql = ZSql & "TipoEtapa ,"
            ZSql = ZSql & "Envasamiento ,"
            ZSql = ZSql & "EquipoII ,"
            ZSql = ZSql & "Desde ,"
            ZSql = ZSql & "Hasta ,"
            ZSql = ZSql & "Lista ,"
            ZSql = ZSql & "Suma1 ,"
            ZSql = ZSql & "Suma2 ,"
            ZSql = ZSql & "Suma3 ,"
            ZSql = ZSql & "Suma4 ,"
            ZSql = ZSql & "Suma5 ,"
            ZSql = ZSql & "Suma6 ,"
            ZSql = ZSql & "Identificacion ,"
            ZSql = ZSql & "NroPedido ,"
            ZSql = ZSql & "FechaVencimiento ,"
            ZSql = ZSql & "OrdFechaVencimiento ,"
            ZSql = ZSql & "Revalida ,"
            ZSql = ZSql & "FechaRevalida ,"
            ZSql = ZSql & "OrdFechaRevalida ,"
            ZSql = ZSql & "MesesRevalida ,"
            ZSql = ZSql & "MarcaVencida ,"
            ZSql = ZSql & "LoteColorante ,"
            ZSql = ZSql & "TipoOri ,"
            ZSql = ZSql & "ImpreVersion ,"
            ZSql = ZSql & "ImpreFechaVersion )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + ZZClave + "',"
            ZSql = ZSql & "'" + Str$(ZZZZHoja) + "',"
            ZSql = ZSql & "'" + Str$(ZZRenglon) + "',"
            ZSql = ZSql & "'" + ZZFecha + "',"
            ZSql = ZSql & "'" + ZZProducto + "',"
            ZSql = ZSql & "'" + Str$(ZZCantidad) + "',"
            ZSql = ZSql & "'" + ZZTipo + "',"
            ZSql = ZSql & "'" + Str$(ZZZZLote) + "',"
            ZSql = ZSql & "'" + ZZArticulo + "',"
            ZSql = ZSql & "'" + ZZTerminado + "',"
            ZSql = ZSql & "'" + Str$(ZZTeorico) + "',"
            ZSql = ZSql & "'" + Str$(ZZReal) + "',"
            ZSql = ZSql & "'" + ZZFechaing + "',"
            ZSql = ZSql & "'" + ZZFechaIngOrd + "',"
            ZSql = ZSql & "'" + ZZWDate + "',"
            ZSql = ZSql & "'" + Str$(ZZWImporte) + "',"
            ZSql = ZSql & "'" + ZZMarca + "',"
            ZSql = ZSql & "'" + Str$(ZZSaldo) + "',"
            ZSql = ZSql & "'" + Str$(ZZZZLote1) + "',"
            ZSql = ZSql & "'" + Str$(ZZCanti1) + "',"
            ZSql = ZSql & "'" + Str$(ZZZZLote2) + "',"
            ZSql = ZSql & "'" + Str$(ZZCanti2) + "',"
            ZSql = ZSql & "'" + Str$(ZZZZLote3) + "',"
            ZSql = ZSql & "'" + Str$(ZZCanti3) + "',"
            ZSql = ZSql & "'" + Str$(ZZCosto1) + "',"
            ZSql = ZSql & "'" + Str$(ZZCosto2) + "',"
            ZSql = ZSql & "'" + Str$(ZZCosto3) + "',"
            ZSql = ZSql & "'" + ZZMarcaAnt + "',"
            ZSql = ZSql & "'" + Str$(ZZSaldoAnt) + "',"
            ZSql = ZSql & "'" + Str$(ZZRealAnt) + "',"
            ZSql = ZSql & "'" + ZZFechaOrd + "',"
            ZSql = ZSql & "'" + ZZEquipo + "',"
            ZSql = ZSql & "'" + ZZMarcaLabora + "',"
            ZSql = ZSql & "'" + ZZEstado + "',"
            ZSql = ZSql & "'" + Str$(ZZVersionI) + "',"
            ZSql = ZSql & "'" + Str$(ZZVersionII) + "',"
            ZSql = ZSql & "'" + Str$(ZZVersionIII) + "',"
            ZSql = ZSql & "'" + ZZEstadoII + "',"
            ZSql = ZSql & "'" + ZZImpreArticulo + "',"
            ZSql = ZSql & "'" + ZZFechaInicio + "',"
            ZSql = ZSql & "'" + ZZHoraInicio + "',"
            ZSql = ZSql & "'" + ZZFechafinal + "',"
            ZSql = ZSql & "'" + ZZHoraFinal + "',"
            ZSql = ZSql & "'" + Str$(ZZPorceDife) + "',"
            ZSql = ZSql & "'" + ZZImpresionI + "',"
            ZSql = ZSql & "'" + ZZImpresionII + "',"
            ZSql = ZSql & "'" + Str$(ZZMotivoDesvio) + "',"
            ZSql = ZSql & "'" + ZZObservaDesvio + "',"
            ZSql = ZSql & "'" + Str$(ZZImpreReal) + "',"
            ZSql = ZSql & "'" + Str$(ZZOperario) + "',"
            ZSql = ZSql & "'" + Str$(ZZEstadoHoja) + "',"
            ZSql = ZSql & "'" + Str$(ZZEtapa) + "',"
            ZSql = ZSql & "'" + ZZFechaInicioEtaoa + "',"
            ZSql = ZSql & "'" + ZZHoraInicioEtapa + "',"
            ZSql = ZSql & "'" + Str$(ZZTimerInicioEtapa) + "',"
            ZSql = ZSql & "'" + ZZAlarma + "',"
            ZSql = ZSql & "'" + Str$(ZZControlI) + "',"
            ZSql = ZSql & "'" + Str$(ZZControlII) + "',"
            ZSql = ZSql & "'" + Str$(ZZDesdeI) + "',"
            ZSql = ZSql & "'" + Str$(ZZHastaI) + "',"
            ZSql = ZSql & "'" + Str$(ZZTiempoI) + "',"
            ZSql = ZSql & "'" + Str$(ZZTiempoII) + "',"
            ZSql = ZSql & "'" + ZZAlarmaI + "',"
            ZSql = ZSql & "'" + ZZAlarmaII + "',"
            ZSql = ZSql & "'" + Str$(ZZAlarmaITiempo) + "',"
            ZSql = ZSql & "'" + Str$(ZZAlarmaITempe) + "',"
            ZSql = ZSql & "'" + Str$(ZZTiempoIII) + "',"
            ZSql = ZSql & "'" + Str$(ZZTemperatura) + "',"
            ZSql = ZSql & "'" + Str$(ZZTipoEtapa) + "',"
            ZSql = ZSql & "'" + ZZEnvasamiento + "',"
            ZSql = ZSql & "'" + Str$(ZZEquipoII) + "',"
            ZSql = ZSql & "'" + Str$(ZZDesde) + "',"
            ZSql = ZSql & "'" + Str$(ZZHasta) + "',"
            ZSql = ZSql & "'" + ZZLista + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma1) + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma2) + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma3) + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma4) + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma5) + "',"
            ZSql = ZSql & "'" + Str$(ZZSuma6) + "',"
            ZSql = ZSql & "'" + ZZIdentificacion + "',"
            ZSql = ZSql & "'" + Str$(ZZNroPedido) + "',"
            ZSql = ZSql & "'" + ZZFechaVencimiento + "',"
            ZSql = ZSql & "'" + ZZOrdFechaVencimiento + "',"
            ZSql = ZSql & "'" + Str$(ZZRevalida) + "',"
            ZSql = ZSql & "'" + ZZFechaRevalida + "',"
            ZSql = ZSql & "'" + ZZOrdFechaRevalida + "',"
            ZSql = ZSql & "'" + Str$(ZZMesesRevalida) + "',"
            ZSql = ZSql & "'" + ZZMarcaVencida + "',"
            ZSql = ZSql & "'" + ZZLoteColorante + "',"
            ZSql = ZSql & "'" + ZZTipoOri + "',"
            ZSql = ZSql & "'" + Str$(ZZImpreVersion) + "',"
            ZSql = ZSql & "'" + ZZImpreFechaVersion + "')"

            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        End If
    
    
    Next Ciclo
    
    Stop
    Stop
    Stop
    
    
    For Ciclo = 1 To WLugar
    
        WHoja = WSalva(Ciclo, 1)
        WProducto = WSalva(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Producto = " + "'" + WProducto + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + WHoja + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Call cmdClose_Click


End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()
    
    Producto.Text = UCase(Producto.Text)
    
    Rem On Error GoTo WError
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha de la hoja de produccion es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
    If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
    
    If Val(NroPedido.Text) <> 0 Then
        
        XEmpresa = Wempresa
        
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZEntra = "N"
        ZZProducto = "PT-5" + Mid$(Producto.Text, 5, 8)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
        ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZSaldo = rstPedido!Cantidad - rstPedido!Facturado
            If ZSaldo >= 0 Then
                ZEntra = "S"
            End If
            rstPedido.Close
        End If
            
        Call Conecta_Empresa
            
        If ZEntra = "N" Then
            m$ = "Nro de Pedido Incorrecto"
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
            Else
            
        WMarca = 0
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WMarca = IIf(IsNull(rstTerminado!Marca), "0", rstTerminado!Marca)
            rstTerminado.Close
        End If
        If WMarca = 0 Then
            m$ = "Se debe informar Nro de Pedido"
            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
    End If
    
    End If
    
    
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        WVersionI = IIf(IsNull(rstHoja!VersionI), "0", rstHoja!VersionI)
        WVersionII = IIf(IsNull(rstHoja!VersionII), "0", rstHoja!VersionII)
        WVersionIII = IIf(IsNull(rstHoja!VersionIII), "0", rstHoja!VersionIII)
    
        rstHoja.Close
        
        If WVersionI = 99 And WVersionII = 99 And WVersionIII = 99 Then
        
            spHoja = "BorrarHoja " + "'" + Hoja.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)
        
                Else
        
            m$ = "Partida ya existente"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        
        End If
        
    End If
    
    Select Case Val(Wempresa)
        Case 1
            Rem If Val(Hoja.Text) > 69999 Or Val(Hoja.Text) < 57600 Then
            If Val(Hoja.Text) > 199999 Or Val(Hoja.Text) < 100000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 100000 y 199999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 2
            If Val(Hoja.Text) > 55999 Or Val(Hoja.Text) < 55300 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 55300 y 55999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 3
            Rem If Val(Hoja.Text) > 99999 Or Val(Hoja.Text) < 82000 Then
            If Val(Hoja.Text) > 299999 Or Val(Hoja.Text) < 200000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 200000 y 299999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 4
            If Val(Hoja.Text) > 19999 Or Val(Hoja.Text) < 11100 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 11100 y 19999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 5
            Rem If Val(Hoja.Text) > 9999 Or Val(Hoja.Text) < 4600 Then
            If Val(Hoja.Text) > 399999 Or Val(Hoja.Text) < 300000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 300000 y 399999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 6
            Rem If Val(Hoja.Text) > 1999 Or Val(Hoja.Text) < 1740 Then
            If Val(Hoja.Text) > 499999 Or Val(Hoja.Text) < 400000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 400000 y 499999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 7
            Rem If Val(Hoja.Text) > 999 Or Val(Hoja.Text) < 7 Then
            If Val(Hoja.Text) > 599999 Or Val(Hoja.Text) < 500000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 500000 y 599999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 8
            If Val(Hoja.Text) > 29999 Or Val(Hoja.Text) < 20800 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 20800 y 29999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 9
            If Val(Hoja.Text) > 30999 Or Val(Hoja.Text) < 30000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 30000 y 30999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 10
            Rem If Val(Hoja.Text) > 999 Or Val(Hoja.Text) < 7 Then
            If Val(Hoja.Text) > 699999 Or Val(Hoja.Text) < 600000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 600000 y 799999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 11
            Rem If Val(Hoja.Text) > 999 Or Val(Hoja.Text) < 7 Then
            If Val(Hoja.Text) > 799999 Or Val(Hoja.Text) < 700000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 600000 y 799999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case Else
    End Select

    If Val(Real.Text) = 0 Then
        Real.Text = "0"
    End If
    
    If Val(Wempresa) = 5 Then
    
        Rem spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstTerminado.RecordCount > 0 Then
        Rem     WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
        Rem     WTeorico = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
        Rem     rstTerminado.Close
        Rem End If
        Rem If WEscrito = 1 Then
        Rem     If Val(Teorico.Text) <> Val(WTeorico) Then
        Rem         Exit Sub
        Rem     End If
        Rem End If
        Rem If WEscrito = 2 Then
        Rem     If Val(Teorico.Text) <> Val(WTeorico) Then
        Rem         Exit Sub
        Rem     End If
        Rem End If
        
            Else
            
        WMetodo = ""
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WMetodo = IIf(IsNull(rstTerminado!Metodo), "", rstTerminado!Metodo)
            rstTerminado.Close
        End If
        If Trim(WMetodo) = "" Then
            m$ = "Se debe informar el metodo de lavado correspondiente"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
            
    End If
    
    ZZLugar = 0
    For A = 1 To 40

        ZZLugar = ZZLugar + 1
        
        Tipo = UCase(Grilla.TextMatrix(ZZLugar, 1))
        Terminado = UCase(Grilla.TextMatrix(ZZLugar, 2))
        Articulo = UCase(Grilla.TextMatrix(ZZLugar, 3))
        Cantidad = Grilla.TextMatrix(ZZLugar, 5)
     
        If Tipo <> AuxiliarIII(ZZLugar, 1) Then
            m$ = "Se detecto un error al verificar la formula de produccion, por favor reinicie la computadora y pruebe nuevamente de ingresarla"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
        If Articulo <> AuxiliarIII(ZZLugar, 2) Then
            m$ = "Se detecto un error al verificar la formula de produccion, por favor reinicie la computadora y pruebe nuevamente de ingresarla"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
        If Terminado <> AuxiliarIII(ZZLugar, 3) Then
            m$ = "Se detecto un error al verificar la formula de produccion, por favor reinicie la computadora y pruebe nuevamente de ingresarla"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
        
        If Val(Cantidad) <> Val(AuxiliarIII(ZZLugar, 4)) Then
            m$ = "Se detecto un error al verificar la formula de produccion, por favor reinicie la computadora y pruebe nuevamente de ingresarla"
            G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
            Exit Sub
        End If
            
    Next A

    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WProducto = Producto.Text
    WTeorico = Teorico.Text
    WReal = Real.Text
    WFechaing = fechaIng.Text
  
    PLote = Hoja.Text
    PTerminado = Producto.Text

    Renglon = 0
    Erase Auxiliar
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    With rstHoja
        .MoveFirst
        Do
            If .EOF = False Then
                Renglon = Renglon + 1
                Auxiliar(Renglon, 1) = rstHoja!Producto
                Auxiliar(Renglon, 2) = rstHoja!Terminado
                Auxiliar(Renglon, 3) = rstHoja!Articulo
                Auxiliar(Renglon, 4) = rstHoja!Cantidad
                Auxiliar(Renglon, 5) = rstHoja!Real
                Auxiliar(Renglon, 6) = rstHoja!Teorico
                Auxiliar(Renglon, 7) = rstHoja!Tipo
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstHoja.Close
    End If
    
    For Da = 1 To Renglon

        Producto = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Real = Auxiliar(Da, 5)
        Teorico = Auxiliar(Da, 6)
        Tipo = Auxiliar(Da, 7)
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                If Real <> 0 Then
                    WEntradas = Str$(rstTerminado!Entradas - Val(Real))
                    WProceso = Str$(rstTerminado!Proceso)
                        Else
                    WProceso = Str$(rstTerminado!Proceso - Val(Teorico))
                    WEntradas = Str$(rstTerminado!Entradas)
                End If
                WDate = Date$
                rstTerminado.Close
                
                XParam = "'" + WCodigo + "','" _
                        + WEntradas + "','" _
                        + WProceso + "','" _
                        + WDate + "'"
                                           
                spTerminado = "ModificaTerminadoHoja " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            End If
                
            Select Case Tipo
                Case "M"
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = rstArticulo!Codigo
                        WSalidas = Str$(rstArticulo!Salidas - Val(Cantidad))
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                                           
                        spArticulo = "ModificaArticuloSalidas " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                                        
                Case "T"
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WCodigo = rstTerminado!Codigo
                        WSalidas = Str$(rstTerminado!Salidas - Val(Cantidad))
                        WDate = Date$
                        rstTerminado.Close
                        
                        XParam = "'" + WCodigo + "','" _
                            + WSalidas + "','" _
                            + WDate + "'"
                                            
                        spTerminado = "ModificaTerminadoSalidas " + XParam
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                Case Else
            End Select
        End If
            
    Next Da
    
    Entra = "S"
    
    For A = 1 To 40
        
        Tipo = UCase(Grilla.TextMatrix(A, 1))
        Terminado = UCase(Grilla.TextMatrix(A, 2))
        Articulo = UCase(Grilla.TextMatrix(A, 3))
        Cantidad = Grilla.TextMatrix(A, 5)
                    
        If Articulo <> "" Then
                
            Select Case Tipo
                Case "T"
                    spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                        WImpre1 = Terminado
                        rstTerminado.Close
                    End If
                Case "M"
                    spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                        WImpre1 = Articulo
                        rstArticulo.Close
                    End If
                Case Else
            End Select
        
            If Val(Cantidad) > WStock Then
                WImpre = Str$(WStock)
                WImpre = Pusing("###,###.###", WImpre)
                m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Entra = "N"
            End If
            
            If Val(Cantidad) = 0 Then
                m$ = "No se informo cantidad para el item " + WImpre1
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Entra = "N"
            End If
            
        End If
            
    Next A
    
    If Entra = "S" Then
    
        spHoja = "BorrarHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)
        
        
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                Call GrabaVencido
            Case Else
        End Select
    
        Renglon = 0
        Erase Auxiliar
        
        Hoja.Text = WHoja
        Fecha.Text = WFecha
        Producto.Text = WProducto
        Teorico.Text = WTeorico
        Real.Text = WReal
        fechaIng.Text = WFechaing
        
        For A = 1 To 40
        
            Tipo = UCase(Grilla.TextMatrix(A, 1))
            Terminado = UCase(Grilla.TextMatrix(A, 2))
            Articulo = UCase(Grilla.TextMatrix(A, 3))
            ImpreArticulo = UCase(Grilla.TextMatrix(A, 4))
            Cantidad = Grilla.TextMatrix(A, 5)
                    
            If Articulo <> "" Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                    
                Auxi1 = Str$(Hoja.Text)
                Call Ceros(Auxi1, 6)
                
                WClave = Auxi1 + Auxi
                WHoja = WHoja
                WRenglon = Str$(Renglon)
                WFecha = WFecha
                WProducto = WProducto
                WTeorico = WTeorico
                WReal = WReal
                WFechaing = WFechaing
                WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
                WTipo = Tipo
                WArticulo = Articulo
                WTerminado = Terminado
                WCantidad = Cantidad
                WLote = ""
                WDate = Date$
                WImporte = ""
                WMarca = ""
                WSaldo = "0"
                ZLugar = A
                WLote1 = ZLote(ZLugar, 1)
                WLote2 = ZLote(ZLugar, 3)
                WLote3 = ZLote(ZLugar, 5)
                WCanti1 = ZLote(ZLugar, 2)
                WCanti2 = ZLote(ZLugar, 4)
                WCanti3 = ZLote(ZLugar, 6)
                WCosto1 = "0"
                WCosto2 = "0"
                WCosto3 = "0"
                
                XParam = "'" + WClave + "','" _
                        + WHoja + "','" _
                        + WRenglon + "','" _
                        + WFecha + "','" _
                        + WProducto + "','" _
                        + WCantidad + "','" _
                        + WTipo + "','" _
                        + WLote + "','" _
                        + WArticulo + "','" _
                        + WTerminado + "','" _
                        + WTeorico + "','" _
                        + WReal + "','" _
                        + WFechaing + "','" _
                        + WFechaingord + "','" _
                        + WDate + "','" _
                        + WImporte + "','" _
                        + WMarca + "','" _
                        + WSaldo + "','" _
                        + WLote1 + "','" + WCanti1 + "','" _
                        + WLote2 + "','" + WCanti2 + "','" _
                        + WLote3 + "','" + WLote3 + "','" _
                        + WCosto1 + "','" _
                        + WCosto2 + "','" _
                        + WCosto3 + "'"
                                       
                spHoja = "AltaHoja " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " ImpreArticulo = " + "'" + ImpreArticulo + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " NroPedido = " + "'" + NroPedido.Text + "',"
                ZSql = ZSql + " Equipo = " + "'" + Equipo.Text + "',"
                ZSql = ZSql + " VersionI = " + "'" + VersionI.Text + "',"
                ZSql = ZSql + " VersionII = " + "'" + VersionII.Text + "',"
                ZSql = ZSql + " VersionIII = " + "'" + VersionIII.Text + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
                Auxiliar(Renglon, 1) = WProducto
                Auxiliar(Renglon, 2) = WTerminado
                Auxiliar(Renglon, 3) = WArticulo
                Auxiliar(Renglon, 4) = WCantidad
                Auxiliar(Renglon, 5) = WReal
                Auxiliar(Renglon, 6) = WTeorico
                Auxiliar(Renglon, 7) = WTipo
                
            End If
            
        Next A
        
        WHoja = Hoja.Text
        WFecha = Fecha.Text
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WHoja + "','" _
                     + WFechaord + "'"
        Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        For Da = 1 To Renglon
    
            Producto = Auxiliar(Da, 1)
            Terminado = Auxiliar(Da, 2)
            Articulo = Auxiliar(Da, 3)
            Cantidad = Auxiliar(Da, 4)
            Real = Auxiliar(Da, 5)
            Teorico = Auxiliar(Da, 6)
            Tipo = Auxiliar(Da, 7)
        
            If Da = 1 Then
        
                spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    If Real <> 0 Then
                        WEntradas = Str$(rstTerminado!Entradas + Val(Real))
                        WProceso = Str$(rstTerminado!Proceso)
                            Else
                        WProceso = Str$(rstTerminado!Proceso + Val(Teorico))
                        WEntradas = Str$(rstTerminado!Entradas)
                    End If
                    WDate = Date$
                    rstTerminado.Close
                    
                    XParam = "'" + WCodigo + "','" _
                            + WEntradas + "','" _
                            + WProceso + "','" _
                            + WDate + "'"
                                           
                    spTerminado = "ModificaTerminadoHoja " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
                
            Select Case Tipo
                    Case "M"
                        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WCodigo = rstArticulo!Codigo
                            WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                            WDate = Date$
                            rstArticulo.Close
                            XParam = "'" + WCodigo + "','" _
                                    + WSalidas + "','" _
                                    + WDate + "'"
                                                
                            spArticulo = "ModificaArticuloSalidas " + XParam
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                                            
                    Case "T"
                        spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            WCodigo = rstTerminado!Codigo
                            WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                            WDate = Date$
                            rstTerminado.Close
                            
                            XParam = "'" + WCodigo + "','" _
                                + WSalidas + "','" _
                                + WDate + "'"
                                            
                            spTerminado = "ModificaTerminadoSalidas " + XParam
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    
                    Case Else
            End Select
        
        Next Da
        
        T$ = "Hoja de Produccion"
        m$ = "Desea Imprimir la Hoja de Produccion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impresion
        End If
        
        Call Limpia_Click
        Hoja.SetFocus
        
        PrgHoja.Hide
        Unload Me
        PrgEti3Auto.Show
        
            Else
            
        Grilla.Col = 1
        Grilla.Row = 1
        
        Hoja.SetFocus
    
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Limpia_Click()

    Graba.Enabled = True
    Reimpresion.Enabled = False

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    observaciones.Text = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    Teorico.Text = ""
    Equipo.Text = ""
    Graba.Enabled = True
    
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    FechaInicio.Text = "  /  /    "
    HoraInicio.Text = ""
    FechaFinal.Text = "  /  /    "
    HoraFinal.Text = ""
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False
    Erase ZLote
    
    Call Limpia_Grilla
    
    Rem With rstHoja
    Rem     .Index = "Clave"
    Rem     Claveven$ = "99999999"
    Rem     .Seek "<=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Hoja.Text = !Hoja + 1
    Rem             Else
    Rem         Hoja.Text = ""
    Rem     End If
    Rem End With
    
    If LeeHoja <> "N" Then
        spHoja = "ListaHojaNumero"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveLast
                Hoja.Text = rstHoja!Hoja + 1
            End With
            rstHoja.Close
                Else
            Hoja.Text = "1"
        End If
    End If
    
    Renglon = 0
    Hoja.SetFocus

End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim CargaEmpresa(12, 2) As String
    
        XEmpresa = Wempresa
        Erase CargaEmpresa
    
        Select Case Val(Wempresa)
            Case 1, 3, 5, 6, 7, 10, 11
                CargaEmpresa(1, 1) = "0001"
                CargaEmpresa(1, 2) = "Empresa01"
                CargaEmpresa(2, 1) = "0003"
                CargaEmpresa(2, 2) = "Empresa03"
                CargaEmpresa(3, 1) = "0005"
                CargaEmpresa(3, 2) = "Empresa05"
                CargaEmpresa(4, 1) = "0006"
                CargaEmpresa(4, 2) = "Empresa06"
                CargaEmpresa(5, 1) = "0007"
                CargaEmpresa(5, 2) = "Empresa07"
                CargaEmpresa(6, 1) = "0010"
                CargaEmpresa(6, 2) = "Empresa10"
                CargaEmpresa(7, 1) = "0011"
                CargaEmpresa(7, 2) = "Empresa11"
            Case Else
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
        End Select
            
        For Cicla = 1 To 7
        
            If CargaEmpresa(Cicla, 1) <> "" Then
        
                Wempresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "Observaciones = " + "'" + observaciones.Text + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + Producto.Text + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next Cicla
    
        Select Case Val(XEmpresa)
            Case 1
                Wempresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 2
                Wempresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 3
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 4
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 5
                Wempresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 6
                Wempresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 7
                Wempresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 8
                Wempresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 9
                Wempresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                Wempresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                Wempresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        m$ = "Actualizacion realizada"
        G% = MsgBox(m$, 0, "Hoja de Produccion")
    End If

End Sub

Private Sub Reimpresion_Click()

    T$ = "Hoja de Produccion"
    m$ = "Desea Imprimir la Hoja de Produccion"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
    
    Grilla.Col = 1
    Grilla.Row = 1

    Hoja.SetFocus
        
End Sub

Private Sub VersionII_Click()

    ZTerminado = Producto.Text
    ZVersion = VersionII.Text
    PrgEspeAuto.Show

End Sub

Private Sub VersionIII_Click()

    ZTerminado = Producto.Text
    ZVersion = VersionIII.Text
    PrgCargaIVAuto.Show
    
End Sub

Private Sub Grilla_dblClick()

    WWLugar = Grilla.Row
    WWTipo = UCase(Grilla.TextMatrix(WWLugar, 1))
    WWTerminado = UCase(Grilla.TextMatrix(WWLugar, 2))
    WWArticulo = UCase(Grilla.TextMatrix(WWLugar, 3))
    WWCantidad = Grilla.TextMatrix(WWLugar, 5)
    
    CargaLote.Visible = True
    
    If WWTipo = "M" Then
        CargaLote.Caption = "Ingreso de Lote:" + WWArticulo
        dada.Caption = "Lote"
            Else
        CargaLote.Caption = "Ingreso de Partida:" + WWTerminado
        dada.Caption = "Partida"
    End If
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    If Val(ZLote(WWLugar, 1)) <> 0 Then
        WLote1.Text = ZLote(WWLugar, 1)
        WCanti1.Text = ZLote(WWLugar, 2)
        WControl1.Locked = False
        WControl1.Text = ""
        WControl1.Locked = True
    End If
    If Val(ZLote(WWLugar, 3)) <> 0 Then
        WLote2.Text = ZLote(WWLugar, 3)
        WCanti2.Text = ZLote(WWLugar, 4)
        WControl2.Locked = False
        WControl2.Text = ""
        WControl2.Locked = True
    End If
    If Val(ZLote(WWLugar, 5)) <> 0 Then
        WLote3.Text = ZLote(WWLugar, 5)
        WCanti3.Text = ZLote(WWLugar, 6)
        WControl3.Locked = False
        WControl3.Text = ""
        WControl3.Locked = True
    End If
    WLote1.SetFocus
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote1.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo1 = 0
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    wdada = rstHoja!Hoja
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo1 >= Val(WCanti1.Text) Or WControla > 0 Then
            WCanti1.Text = Pusing("###,###.###", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            If WWTipo = "M" Then
                m$ = WWArticulo + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WWTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote1.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote2.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo2 = 0
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + WLote2.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo2 >= Val(WCanti2.Text) Or WControla > 0 Then
            WCanti2.Text = Pusing("###,###.###", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
            WLote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            If WWTipo = "M" Then
                m$ = WWArticulo + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WWTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote2.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
            
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If Val(WLote3.Text) = 0 Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    Call Alta_Vector
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WSaldo3 = 0
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
    
        If WSaldo3 >= Val(WCanti3.Text) Or WControla > 0 Then
            WCanti3.Text = Pusing("###,###.###", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
            Call Verifica_Lote
            If WEstado = "S" Then
                Call Alta_Vector
            End If
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            If WWTipo = "M" Then
                m$ = WWArticulo + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
                    Else
                m$ = WWTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            WLote3.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub





Sub Indica()

    Select Case XIndice
        Case 0
            Producto.SetFocus
        Case 1, 2
        Case Else
    End Select

End Sub

Private Sub Form_Load()

    Call Limpia_Grilla
    
    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False

    Hoja.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    observaciones.Text = ""
    fechaIng.Text = "  /  /    "
    Real.Text = ""
    NroPedido.Text = ""
    Teorico.Text = ""
    Graba.Enabled = True
    Reimpresion.Enabled = False
    If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
        LabelPedido.Visible = True
        NroPedido.Visible = True
    End If
    
    Equipo.Text = ""
    VersionI.Text = ""
    VersionII.Text = ""
    VersionIII.Text = ""
    FechaInicio.Text = "  /  /    "
    HoraInicio.Text = ""
    FechaFinal.Text = "  /  /    "
    HoraFinal.Text = ""
    
    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            Hoja.Text = rstHoja!Hoja + 1
        End With
        rstHoja.Close
            Else
        Hoja.Text = "1"
    End If
    
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgHoja.Caption = "Ingreso de Hoja de produccion :  " + !Nombre
        End If
    End With
    
    Muestra.ColWidth(0) = 150
    Muestra.ColWidth(1) = 400
    Muestra.ColWidth(2) = 800
    Muestra.ColWidth(3) = 800
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Partida"
    
    Muestra.Col = 3
    Muestra.Text = "Stock"
    
    Muestra1.ColWidth(0) = 150
    Muestra1.ColWidth(1) = 800
    Muestra1.ColWidth(2) = 800
    Muestra1.ColWidth(3) = 1000
    
    Muestra1.Row = 0
    
    Muestra1.Col = 1
    Muestra1.Text = "Cliente"
    
    Muestra1.Col = 3
    Muestra1.Text = "Fecha"
    
    Muestra1.Col = 2
    Muestra1.Text = "Cantidad"
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Grilla
    
    Renglon = 0
    Erase Auxiliar
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Renglon = Renglon + 1
                
                    Grilla.TextMatrix(Renglon, 1) = rstHoja!Tipo
                    Grilla.TextMatrix(Renglon, 2) = rstHoja!Terminado
                    Grilla.TextMatrix(Renglon, 3) = rstHoja!Articulo
                    If rstHoja!Cantidad >= 1 Then
                        Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", rstHoja!Cantidad)
                            Else
                        Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.###", rstHoja!Cantidad)
                    End If
                
                    Auxiliar(Renglon, 1) = rstHoja!Tipo
                    Auxiliar(Renglon, 2) = rstHoja!Terminado
                    Auxiliar(Renglon, 3) = rstHoja!Articulo
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
            
        Tipo = Auxiliar(Renglon, 1)
        Auxi1 = Auxiliar(Renglon, 2)
        Auxi2 = Auxiliar(Renglon, 3)
                
        Select Case Tipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Grilla.TextMatrix(Renglon, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Grilla.TextMatrix(Renglon, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            Case Else
        End Select
    Next Da

    If Val(Real.Text) <> 0 Then
        Graba.Enabled = False
    End If
  
End Sub

Private Sub Hoja_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
        Entra = "N"
        Existe = "N"
        spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            Existe = "S"
            Fecha.Text = rstHoja!Fecha
            Real.Text = Str$(rstHoja!Real)
            Teorico.Text = Str$(rstHoja!Teorico)
            fechaIng.Text = rstHoja!fechaIng
            Producto.Text = rstHoja!Producto
            NroPedido.Text = IIf(IsNull(rstHoja!NroPedido), "", rstHoja!NroPedido)
            Equipo.Text = IIf(IsNull(rstHoja!Equipo), "", rstHoja!Equipo)
            VersionI.Text = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
            VersionII.Text = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
            VersionIII.Text = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
            FechaFinal.Text = IIf(IsNull(rstHoja!FechaFinal), "  /  /    ", rstHoja!FechaFinal)
            HoraFinal.Text = IIf(IsNull(rstHoja!HoraFinal), "", rstHoja!HoraFinal)
            FechaInicio.Text = IIf(IsNull(rstHoja!FechaInicio), "  /  /    ", rstHoja!FechaInicio)
            HoraInicio.Text = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
            HoraInicio.Text = Trim(HoraInicio.Text)
            HoraFinal.Text = Trim(HoraFinal.Text)
            rstHoja.Close
        
        End If
        
        If Existe = "S" Then
            
            If Val(VersionI.Text) = 99 And Val(VersionII.Text) = 99 And Val(VersionIII.Text) = 99 Then
            
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
            
                    If Val(Wempresa) = 1 Or Val(Wempresa) = 2 Or Val(Wempresa) = 3 Or Val(Wempresa) = 4 Or Val(Wempresa) = 7 Then
            
                        WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                        WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                        WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                
                        If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                            m$ = "El Producto Terminado no se encuentra autorizado para la Produccion"
                            If WEstadoI = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                            End If
                            If WEstadoII = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                            End If
                            If WEstadoIII = "N" Then
                                m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                            End If
                            ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                            Exit Sub
                        End If
                
                    End If
                
                    VersionI.Text = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
                    VersionII.Text = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
                    VersionIII.Text = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
            
                    Producto.Text = rstTerminado!Codigo
                    DesProducto.Caption = rstTerminado!Descripcion
                    observaciones.Text = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
                    rstTerminado.Close
                    
                    Call Calcula_stock
                    Call Lee_Composicion
                    
                
                End If
            
                    Else
            
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Producto.Text = rstTerminado!Codigo
                    DesProducto.Caption = rstTerminado!Descripcion
                    observaciones.Text = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
                    rstTerminado.Close
                End If
                Graba.Enabled = False
                Reimpresion.Enabled = True
                Call Proceso_Click
            
            End If
                
                Else
                    
            Graba.Enabled = True
            Reimpresion.Enabled = False
            Existe = "N"
                    
            WHoja = Hoja.Text
            LeeHoja = "N"
            Call Limpia_Click
            LeeHoja = "S"
            Hoja.Text = WHoja
            Producto.SetFocus
                
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    Rem If KeyAscii = 13 Then
    Rem     Call Valida_fecha(Fecha.Text, Auxi)
    Rem     If Auxi = "S" Then
    Rem        Producto.SetFocus
    Rem            Else
    Rem        Fecha.SetFocus
    Rem    End If
    Rem End If
End Sub

Private Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
                If Val(Wempresa) = 1 Or Val(Wempresa) = 2 Or Val(Wempresa) = 3 Or Val(Wempresa) = 4 Or Val(Wempresa) = 7 Then
            
                    WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                    WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                    WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
                
                    If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                        m$ = "El Producto Terminado no se encuentra autorizado para la Produccion"
                        If WEstadoI = "N" Then
                            m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                        End If
                        If WEstadoII = "N" Then
                            m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                        End If
                        If WEstadoIII = "N" Then
                            m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                        End If
                        ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        Exit Sub
                    End If
                
                End If
                
                VersionI.Text = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
                VersionII.Text = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
                VersionIII.Text = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
                
                WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
            
                Producto.Text = rstTerminado!Codigo
                DesProducto.Caption = rstTerminado!Descripcion
                observaciones.Text = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
                observaciones.Text = Trim(observaciones.Text)
                rstTerminado.Close
                Call Calcula_stock
                If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
                    NroPedido.SetFocus
                        Else
                    Teorico.SetFocus
                End If
                
                If Val(Wempresa) = 5 And (WEscrito = 1 Or WEscrito = 2) Then
                
                    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        ZZFabrica = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
                        ZZFabricaII = IIf(IsNull(rstTerminado!FabricaII), "0", rstTerminado!FabricaII)
                        ZZFabricaIII = IIf(IsNull(rstTerminado!FabricaIII), "0", rstTerminado!FabricaIII)
                        Fabrica.Text = Str$(ZZFabrica)
                        FabricaII.Text = Str$(ZZFabricaII)
                        FabricaIII.Text = Str$(ZZFabricaIII)
                        rstTerminado.Close
                    End If
                    
                    PantaTeorico.Visible = True
                    Fabrica.SetFocus
                        
                        Else
                        
                    If Trim(observaciones.Text) <> "" Then
                        ImpreObserva.Text = observaciones.Text
                        PantaObserva.Height = 3375
                        PantaObserva.Left = 480
                        PantaObserva.Top = 1200
                        PantaObserva.Width = 8655
                        PantaObserva.Visible = True
                    End If
                        
                    
                End If
                
            End If
        End If
    End If
End Sub



Private Sub Fabrica_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(FabricaII.Text) <> 0 And Val(FabricaIII.Text) <> 0 Then
            
            If Val(Fabrica.Text) < Val(FabricaII.Text) Or Val(Fabrica.Text) > Val(FabricaIII.Text) Then
                Exit Sub
            End If
        End If
    
        PantaTeorico.Visible = False
        
        spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
        
            If Val(Wempresa) = 1 Or Val(Wempresa) = 2 Or Val(Wempresa) = 3 Or Val(Wempresa) = 4 Or Val(Wempresa) = 7 Then
        
                WEstadoI = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
                WEstadoII = IIf(IsNull(rstTerminado!EstadoI), "", rstTerminado!EstadoI)
                WEstadoIII = IIf(IsNull(rstTerminado!EstadoII), "", rstTerminado!EstadoII)
            
                If WEstadoI = "N" Or WEstadoII = "N" Or WEstadoIII = "N" Then
                    m$ = "El Producto Terminado no se encuentra autorizado para la Produccion"
                    If WEstadoI = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada la formula)"
                    End If
                    If WEstadoII = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada los procesos)"
                    End If
                    If WEstadoIII = "N" Then
                        m$ = m$ + Chr$(13) + "(No se encuentra habilitada las especificaciones)"
                    End If
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                    Exit Sub
                End If
            
            End If
            
            VersionI.Text = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
            VersionII.Text = IIf(IsNull(rstTerminado!VersionI), "", rstTerminado!VersionI)
            VersionIII.Text = IIf(IsNull(rstTerminado!VersionII), "", rstTerminado!VersionII)
            
            WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
        
            Producto.Text = rstTerminado!Codigo
            DesProducto.Caption = rstTerminado!Descripcion
            observaciones.Text = IIf(IsNull(rstTerminado!observaciones), "", rstTerminado!observaciones)
            rstTerminado.Close
            Call Calcula_stock
            If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
                NroPedido.SetFocus
                    Else
                Teorico.SetFocus
            End If
            
            If Val(Wempresa) = 5 And (WEscrito = 1 Or WEscrito = 2) Then
            
                Erase ZComparaI
                ZRenglonI = 0

                spComposicion = "ConsultaComposicionProducto " + "'" + Producto.Text + "'"
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstComposicion.RecordCount > 0 Then
                    With rstComposicion
                        .MoveFirst
                        Do
                            If .EOF = False Then
    
                                ZZEntraCompo = "S"
                                
                                If rstComposicion!Tipo = "M" Then
                                    If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                                        ZZEntraCompo = "N"
                                    End If
                                End If
                                
                                If ZZEntraCompo = "S" Then
                                
                                    ZRenglonI = ZRenglonI + 1
                                    If rstComposicion!Tipo = "M" Then
                                        ZComparaI(ZRenglonI, 1) = rstComposicion!Articulo1
                                            Else
                                        ZComparaI(ZRenglonI, 1) = rstComposicion!Articulo2
                                    End If
                                    ZComparaI(ZRenglonI, 2) = Str$(rstComposicion!Cantidad)
                                    
                                End If
            
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstComposicion.Close
                End If
                
                Erase ZComparaII
                ZRenglonII = 0
                    
                For Cicla = 1 To 99

                    ZPaso = Str$(Cicla)
    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM CargaIII"
                    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Producto.Text + "'"
                    ZSql = ZSql + " and CargaIII.Paso = " + "'" + ZPaso + "'"
                    ZSql = ZSql + " Order by CargaIII.Clave"
                    rscargaIII = ZSql
                    Set rstCargaIII = db.OpenRecordset(rscargaIII, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCargaIII.RecordCount > 0 Then
                        With rstCargaIII
                            .MoveFirst
                            Do
                                If .EOF = False Then
                                        
                                    If rstCargaIII!Cantidad <> 0 Then
                                        ZRenglonII = ZRenglonII + 1
                                        If Trim(rstCargaIII!Articulo) <> "" Then
                                            ZComparaII(ZRenglonII, 1) = rstCargaIII!Articulo
                                                Else
                                            ZComparaII(ZRenglonII, 1) = rstCargaIII!PTerminado
                                        End If
                                        ZComparaII(ZRenglonII, 2) = Str$(rstCargaIII!Cantidad)
                                    End If
                    
                                    .MoveNext
                                        Else
                                    Exit Do
                                End If
                            Loop
                        End With
                        rstCargaIII.Close
                    End If
                    
                Next Cicla
                
                If ZRenglonI <> ZRenglonII Then
                    m$ = "El procedimiento de fabricacion no es compatible con la formula del producto"
                    G% = MsgBox(m$, 0, "Modificacion de Materia Prima")
                    Exit Sub
                End If
                
                For Cicla = 1 To ZRenglonI
                    If ZComparaI(Cicla, 1) <> ZComparaII(Cicla, 1) Then
                        m$ = "El procedimiento de fabricacion no es compatible con la formula del producto"
                        G% = MsgBox(m$, 0, "Modificacion de Materia Prima")
                        Exit Sub
                    End If
                    If ZComparaI(Cicla, 2) <> ZComparaII(Cicla, 2) Then
                        m$ = "El procedimiento de fabricacion no es compatible con la formula del producto"
                        G% = MsgBox(m$, 0, "Modificacion de Materia Prima")
                        Exit Sub
                    End If
                Next Cicla
                
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZTeorico = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
                    Teorico.Text = Str$(ZZTeorico)
                    rstTerminado.Close
                End If
                
                Teorico.Text = Fabrica.Text
                
                Call Lee_Composicion
            
            End If
            
            XEmpresa = Wempresa
            
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                Wempresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Else
                Wempresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM BajaLote"
            ZSql = ZSql + " Where BajaLote.Destino = " + "'" + Producto.Text + "'"
            ZSql = ZSql + " and BajaLote.Estado = 0"
            ZSql = ZSql + " Order by Numero"
            spBajaLote = ZSql
            Set rstbajalote = db.OpenRecordset(spBajaLote, dbOpenSnapshot, dbSQLPassThrough)
            If rstbajalote.RecordCount > 0 Then
                With rstbajalote
                
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                        
                            If rstbajalote!Articulo = "  -   -   " Then
                                m$ = "SE PUEDE UTILIZAR EL PRODUCTO TERMINADO = " _
                                + rstbajalote!Terminado + Chr$(13) _
                                + "EN LA ELABORACION SEGUN LA SIGUIENTES ESPECIFICACIONES : " + Chr$(13) _
                                + rstbajalote!observaciones + Chr$(13) _
                                + "Planta : " + Str$(rstbajalote!Planta) + Chr$(13) _
                                + "Lote : " + Str$(rstbajalote!Partida)
                                ca% = MsgBox(m$, 64, "Observaciones de Laboratorio")
                                    Else
                                m$ = "SE PUEDE UTILIZAR LA MATERIA PRIMA = " _
                                + rstbajalote!Articulo + Chr$(13) _
                                + "EN LA ELABORACION SEGUN LA SIGUIENTES ESPECIFICACIONES : " + Chr$(13) _
                                + rstbajalote!observaciones + Chr$(13) _
                                + "Planta : " + Str$(rstbajalote!Planta) + Chr$(13) _
                                + "Lote : " + Str$(rstbajalote!Partida)
                                ca% = MsgBox(m$, 64, "Observaciones de Laboratorio")
                            End If
                            
                            .MoveNext
                            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                        Loop
                    End If
                    
                End With
                rstbajalote.Close
            End If
            
            Wempresa = XEmpresa
            Call Conecta_Empresa
            
            If Trim(observaciones.Text) <> "" Then
                ImpreObserva.Text = observaciones.Text
                PantaObserva.Height = 3375
                PantaObserva.Left = 480
                PantaObserva.Top = 1200
                PantaObserva.Width = 8655
                PantaObserva.Visible = True
            End If
            
        End If
    End If
End Sub


Private Sub NroPedido_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Wempresa) = 2 Or Val(Wempresa) = 4 Then
        
            XEmpresa = Wempresa
        
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZEntra = "N"
            ZZProducto = "PT-5" + Mid$(Producto.Text, 5, 8)
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Pedido = " + "'" + NroPedido.Text + "'"
            ZSql = ZSql + " and Pedido.Terminado = " + "'" + ZZProducto + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZSaldo = rstPedido!Cantidad - rstPedido!Facturado
                If ZSaldo >= 0 Then
                    ZEntra = "S"
                End If
                rstPedido.Close
            End If
            
            Call Conecta_Empresa
            
            If ZEntra = "N" Then
                m$ = "Nro de Pedido Incorrecto"
                ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                NroPedido.SetFocus
                    Else
                Teorico.SetFocus
            End If
            
                Else
                
            Teorico.SetFocus
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Teorico_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem If Val(WEmpresa) = 5 Then
        Rem     spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
        Rem     Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstTerminado.RecordCount > 0 Then
        Rem         WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
        Rem         WTeorico = IIf(IsNull(rstTerminado!Fabrica), "0", rstTerminado!Fabrica)
        Rem         rstTerminado.Close
        Rem     End If
        Rem     If WEscrito = 1 Then
        Rem         Teorico.Text = Str$(WTeorico)
        Rem         Exit Sub
        Rem     End If
        Rem     If WEscrito = 2 Then
        Rem         Teorico.Text = Str$(WTeorico)
        Rem         Exit Sub
        Rem     End If
        Rem End If
        Equipo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Equipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Existe = "S" Then
            spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
                WVersionI = IIf(IsNull(rstHoja!VersionI), "0", rstHoja!VersionI)
                WVersionII = IIf(IsNull(rstHoja!VersionII), "0", rstHoja!VersionII)
                WVersionIII = IIf(IsNull(rstHoja!VersionIII), "0", rstHoja!VersionIII)
                rstHoja.Close
                If WVersionI = 99 And WVersionII = 99 And WVersionIII = 99 Then
                    Call Lee_Composicion
                End If
            End If
                Else
            If Val(Wempresa) = 5 Then
                spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                    rstTerminado.Close
                End If
                If WEscrito = 1 Then
                    Exit Sub
                End If
                If WEscrito = 2 Then
                    Exit Sub
                End If
            End If
            Call Lee_Composicion
        End If
    End If
End Sub

Private Sub Lee_Composicion()

    Call Limpia_Grilla

    Erase Auxiliar
    Erase AuxiliarII
    Erase AuxiliarIII
    
    
    
    
    
    
    Renglon = 0
    ZZCantidad = 0
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Producto.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
        
                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
                                        
                        Renglon = Renglon + 1
                        
                        Grilla.TextMatrix(Renglon, 1) = rstComposicion!Tipo
                    
                        If rstComposicion!Articulo1 = "  -   -  " Then
                            Grilla.TextMatrix(Renglon, 3) = "  -   -   "
                            Auxi1 = "  -   -   "
                                Else
                            Grilla.TextMatrix(Renglon, 3) = rstComposicion!Articulo1
                            Auxi1 = rstComposicion!Articulo1
                        End If
                    
                        Grilla.TextMatrix(Renglon, 2) = rstComposicion!Articulo2
                        Auxi2 = rstComposicion!Articulo2
                    
                        Cantidad = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                        ZZCantidad = rstComposicion!Cantidad
                    
                        If Val(Cantidad) > 1 Then
                            ZZCantidadHoja = Val(Cantidad)
                            Call Redondeo(ZZCantidadHoja)
                            Cantidad = Str$(ZZCantidadHoja)
                            Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", Cantidad)
                                Else
                            Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.###", Cantidad)
                        End If
                    
                        Auxiliar(Renglon, 1) = rstComposicion!Tipo
                        Auxiliar(Renglon, 2) = Auxi1
                        Auxiliar(Renglon, 3) = Auxi2
                        If Val(Cantidad) > 1 Then
                            ZZCantidadHoja = Val(Cantidad)
                            Call Redondeo(ZZCantidadHoja)
                            Cantidad = Str$(ZZCantidadHoja)
                            Auxiliar(Renglon, 4) = Pusing("###,###.##", Cantidad)
                                Else
                            Auxiliar(Renglon, 4) = Pusing("###,###.###", Cantidad)
                        End If
                        
                        AuxiliarIII(Renglon, 1) = rstComposicion!Tipo
                        AuxiliarIII(Renglon, 2) = Auxi1
                        AuxiliarIII(Renglon, 3) = Auxi2
                        If Val(Cantidad) > 1 Then
                            ZZCantidadHoja = Val(Cantidad)
                            Call Redondeo(ZZCantidadHoja)
                            Cantidad = Str$(ZZCantidadHoja)
                            AuxiliarIII(Renglon, 4) = Pusing("###,###.##", Cantidad)
                                Else
                            AuxiliarIII(Renglon, 4) = Pusing("###,###.###", Cantidad)
                        End If
                        
                        AuxiliarII(Renglon) = "N"
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    Erase ZCargaDesvio
    PasaDesvio = "N"
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
            
        Tipo = Auxiliar(Renglon, 1)
        Auxi2 = Auxiliar(Renglon, 2)
        Auxi1 = Auxiliar(Renglon, 3)
        XCantidad = Val(Auxiliar(Renglon, 4))
        
        WStock = 0
        
        Select Case Tipo
            Case "T"
                WImpre1 = Auxi1
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Grilla.TextMatrix(Renglon, 4) = rstTerminado!Descripcion
                    WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                End If
                ZZDesvio = 0
                
                ZZBloqueo = 0
                If Left$(Producto.Text, 2) = "PT" Then
                    ZZCodigoBloqueo = Auxi1
                    ZZTipoBloqueo = "T"
                    Call Calcula_Bloqueo
                    If ZZBloqueo > 0 Then
                        m$ = "Existe el producto " + Auxi1 + " la cantidad de : " + Str$(ZZBloqueo) + " Kgs. Bloqueados" + Chr$(13) + "Comuniquese con el laboratorio para su liberacion"
                        ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        WStock = WStock - ZZBloqueo
                    End If
                End If
                
                
            Case "M"
                WImpre1 = Auxi2
                WStock = 0
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Grilla.TextMatrix(Renglon, 4) = rstArticulo!Descripcion
                    WStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                End If
                
                WWArticulo = Auxi2
                WWVencido = 0
                Call Calcula_Stock_Vencido
                If WWVencido > 0 Then
                    m$ = "Existe la materia prima " + Auxi2 + " la cantidad de : " + Str$(WWVencido) + " Kgs. vencidos." + Chr$(13) + "Comuniquese con el laboratorio para su revalida"
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                    WStock = WStock - WWVencido
                End If
                
                Rem by nan me fijo mono
                Rem si esta vencida en pta5 no se deja hacer hoja
                If Wempresa = "0007" And WRenglon = 1 And ZZCantidad = 1 And WWVencido > 0 Then
                    WStock = 0
                End If
                
                Rem SACO NAN
                ZZDesvio = 0
                ZZArticulo = Auxi2
                Call Calcula_Desvio
                
                If Left$(Producto.Text, 2) = "PT" Then
                    ZZBloqueo = 0
                    ZZCodigoBloqueo = Auxi2
                    ZZTipoBloqueo = "M"
                    Call Calcula_Bloqueo
                    If ZZBloqueo > 0 Then
                        m$ = "Existe la materia prima " + Auxi2 + " la cantidad de : " + Str$(ZZBloqueo) + " Kgs. Bloqueados" + Chr$(13) + "Comuniquese con el laboratorio para su liberacion"
                        ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                        WStock = WStock - ZZBloqueo
                    End If
                End If
                
            Case Else
        End Select
        
        If ZZDesvio = 0 Then
        
            If XCantidad <= WStock Then
                If XCantidad >= 1 Then
                    Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(XCantidad))
                        Else
                    Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.###", Str$(XCantidad))
                End If
                    Else
                WImpre = Str$(WStock)
                WImpre = Pusing("###,###.##", WImpre)
                m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
                ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Grilla.TextMatrix(Renglon, 5) = "0"
            End If
            
                Else
                
            If XCantidad <= WStock - ZZDesvio Then
            
                T$ = "Hoja de Produccion"
                m$ = "Materia Prima " + ZZArticulo + " - " + "Existen partidas aprobadas con desvio. Desea utilizarlas"
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    AuxiliarII(Renglon) = "S"
                    PasaDesvio = "S"
                    ZCargaDesvio(Renglon, 1) = ZZArticulo
                    ZCargaDesvio(Renglon, 2) = Str$(XCantidad)
                    ZCargaDesvio(Renglon, 3) = "S"
                End If
                If XCantidad >= 1 Then
                    Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(XCantidad))
                        Else
                    Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.###", Str$(XCantidad))
                End If
                
                    Else
                    
                If XCantidad <= WStock Then
                
                    T$ = "Hoja de Produccion"
                    m$ = "Materia Prima " + ZZArticulo + " - " + "Atencion : Alcanza el stock necesario para la pruduccion teniendo en cuenta partidas aprobadas por desvio. Desea utilizarlas"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                        
                        If XCantidad >= 1 Then
                            Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(XCantidad))
                                Else
                            Grilla.TextMatrix(Renglon, 5) = Pusing("###,###.###", Str$(XCantidad))
                        End If
                        
                        AuxiliarII(Renglon) = "S"
                        PasaDesvio = "S"
                        ZCargaDesvio(Renglon, 1) = ZZArticulo
                        ZCargaDesvio(Renglon, 2) = Str$(XCantidad)
                        ZCargaDesvio(Renglon, 3) = "S"
                        
                            Else
                            
                        Grilla.TextMatrix(Renglon, 5) = "0"
                        
                    End If
                    
                        Else
                        
                    WImpre = Str$(WStock)
                    WImpre = Pusing("###,###.##", WImpre)
                    m$ = "No existe stock suficiente del item " + WImpre1 + " Stock: " + WImpre + " Kgs."
                    ca% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                    Grilla.TextMatrix(Renglon, 5) = "0"
                    
                End If
                
            End If
            
        End If
        
    Next Da
    
    If PasaDesvio = "S" Then
        PrgHojaDesvio.Show
    End If

End Sub

Sub Impresion()

    XCodigo = Val(Mid$(Producto.Text, 4, 5))
    XTipoPro = ""
    If Val(Wempresa) = 1 Then
        If XCodigo >= 0 And XCodigo <= 999 Then
            XTipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                XTipoPro = "CO"
                    Else
                XTipoPro = ""
            End If
        End If
    End If
    
    Rem call Lee

    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 2, 3, 4, 7, 8, 9
            Sql1 = "DELETE ImpreHoja"
            spImpreHoja = Sql1
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Sql1 = "DELETE ImpreHojaII"
            spImpreHojaII = Sql1
            Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
            
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
            
            Call LeeDatos
            
            Call Conecta_Empresa
        
            WHoja = Hoja.Text
            WFecha = Fecha.Text
            WCodigo1 = Left$(Producto.Text, 2)
            WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
            Select Case Val(Wempresa)
                Case 1
                    WMaquina = "I"
                Case 2
                    WMaquina = "I"
                Case 3
                    WMaquina = "II"
                Case 4
                    WMaquina = "II"
                Case 5
                    WMaquina = "III"
                Case 6
                    WMaquina = "IV"
                Case 7
                    WMaquina = "V"
                Case 8
                    WMaquina = "V"
                Case 9
                    WMaquina = "VI"
                Case 10
                    WMaquina = "VI"
                Case 11
                    WMaquina = "VII"
                Case Else
            End Select
            WTeorico = Teorico.Text
            
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZMetodo = IIf(IsNull(rstTerminado!Metodo), "0", rstTerminado!Metodo)
                ZEfluentes = IIf(IsNull(rstTerminado!Efluentes), "0", rstTerminado!Efluentes)
                ZVersionI = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
                ZVersionII = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
                ZVersionIII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
                rstTerminado.Close
            End If
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Efluentes"
            ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + ZEfluentes + "'"
            spEfluentes = ZSql
            Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
            If rstEfluentes.RecordCount > 0 Then
                ZDesEfluentesI = rstEfluentes!Descripcion
                ZDesEfluentesII = rstEfluentes!DescripcionII
                rstEfluentes.Close
            End If

            Linea = 0
            LineaII = 0
            ZZLugar = 0
        
            For A = 1 To 40
        
                ZZLugar = ZZLugar + 1
            
                Tipo = UCase(Grilla.TextMatrix(A, 1))
                Terminado = UCase(Grilla.TextMatrix(A, 2))
                Articulo = UCase(Grilla.TextMatrix(A, 3))
                Cantidad = Grilla.TextMatrix(A, 5)
                ZZZObserva = Grilla.TextMatrix(A, 6)
             
                If Tipo = "M" Then
                
                    Erase Impre
                    Xlugar = 0
                    XCanti = Val(Cantidad)
                    
                    Impre(1, 1) = Val(ZLote(A, 1))
                    Impre(1, 2) = Val(ZLote(A, 2))
                    Impre(2, 1) = Val(ZLote(A, 3))
                    Impre(2, 2) = Val(ZLote(A, 4))
                    Impre(3, 1) = Val(ZLote(A, 5))
                    Impre(3, 2) = Val(ZLote(A, 6))
                    
                    If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                    
                        Impre(1, 1) = 0
                        Impre(1, 2) = 0
                        Impre(2, 1) = 0
                        Impre(2, 2) = 0
                        Impre(3, 1) = 0
                        Impre(3, 2) = 0
                        Impre(4, 1) = 0
                        Impre(4, 2) = 0
                        Impre(5, 1) = 0
                        Impre(5, 2) = 0
                        Impre(6, 1) = 0
                        Impre(6, 2) = 0
                        Impre(7, 1) = 0
                        Impre(7, 2) = 0
                        Impre(8, 1) = 0
                        Impre(8, 2) = 0
                        Impre(9, 1) = 0
                        Impre(9, 2) = 0
                        Impre(10, 1) = 0
                        Impre(10, 2) = 0
                        
                        If AuxiliarII(ZZLugar) = "S" Then

                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
                            ZSql = ZSql + " Order by Laudo.Laudo"
                            spLaudo = ZSql
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                        
                                With rstLaudo
                                    .MoveFirst
                                    If .NoMatch = False Then
                                        Do
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                                            
                                            WMarcaVencida = IIf(IsNull(rstLaudo!MarcaVencida), "", rstLaudo!MarcaVencida)
                                            WMarcaEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                                            
                                            If rstLaudo!Saldo <> 0 And Trim(WMarcaVencida) = "" And Trim(WMarcaEstado) <> "N" Then
                                            
                                                If rstLaudo!Articulo = Articulo Then
                                                
                                                    WLaudo = rstLaudo!Laudo
                                                    ZEntra = "N"
                                                
                                                    If WLaudo = Val(ZCargaDesvio(ZZLugar, 4)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 5)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 6)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 7)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 8)) Then
                                                
                                                        If WLaudo >= 190000 And WLaudo <= 194999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 990000 And WLaudo <= 994999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 290000 And WLaudo <= 294999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 390000 And WLaudo <= 394999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 490000 And WLaudo <= 494999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 590000 And WLaudo <= 594999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 690000 And WLaudo <= 694999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 790000 And WLaudo <= 794999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 890000 And WLaudo <= 894999 Then
                                                            ZEntra = "S"
                                                        End If
                                                
                                                        If ZEntra = "S" Then
                                                        
                                                            XXCanti = 0
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 4)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 9))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 5)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 10))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 6)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 11))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 7)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 12))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 8)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 13))
                                                            End If
                                                    
                                                            If Xlugar < 10 And XCanti > 0 Then
                                                                If XXCanti > 0 Then
                                                                
                                                                    Xlugar = Xlugar + 1
                                                                    If rstLaudo!Saldo > XXCanti Then
                                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                                        Impre(Xlugar, 2) = XXCanti
                                                                        XCanti = XCanti - XXCanti
                                                                            Else
                                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                                        Impre(Xlugar, 2) = rstLaudo!Saldo
                                                                        XCanti = XCanti - rstLaudo!Saldo
                                                                    End If
                                                                    
                                                                        Else
                                                                        
                                                                    Xlugar = Xlugar + 1
                                                                    If rstLaudo!Saldo > XCanti Then
                                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                                        Impre(Xlugar, 2) = XCanti
                                                                        XCanti = 0
                                                                            Else
                                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                                        Impre(Xlugar, 2) = rstLaudo!Saldo
                                                                        XCanti = XCanti - rstLaudo!Saldo
                                                                    End If
                                                                    
                                                                End If
                                                            End If
                                                        
                                                        End If
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
                    
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
                            ZSql = ZSql + " Order by Guia.Codigo"
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                           
                                With rstMovguia
                           
                                    .MoveFirst
                           
                                    If .NoMatch = False Then
                                        Do
                           
                                            If .EOF = True Then
                                                Exit Do
                                            End If
                               
                                            WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                            WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                            
                                            If rstMovguia!Saldo <> 0 And Trim(WMarcaVencida) = "" And Trim(WMarcaEstado) <> "N" Then
                                            
                                                If rstMovguia!Articulo = Articulo Then
                                                
                                                    WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                                                    ZEntra = "N"
                                                    
                                                    If WLaudo = Val(ZCargaDesvio(ZZLugar, 4)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 5)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 6)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 7)) Or WLaudo = Val(ZCargaDesvio(ZZLugar, 8)) Then
                                                    
                                                        If WLaudo >= 190000 And WLaudo <= 194999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 990000 And WLaudo <= 994999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 290000 And WLaudo <= 294999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 390000 And WLaudo <= 394999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 490000 And WLaudo <= 494999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 590000 And WLaudo <= 594999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 690000 And WLaudo <= 694999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 790000 And WLaudo <= 794999 Then
                                                            ZEntra = "S"
                                                        End If
                                                        If WLaudo >= 890000 And WLaudo <= 894999 Then
                                                            ZEntra = "S"
                                                        End If
                                                    
                                                        If ZEntra = "S" Then
                                                        
                                                            XXCanti = 0
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 4)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 9))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 5)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 10))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 6)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 11))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 7)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 12))
                                                            End If
                                                            If WLaudo = Val(ZCargaDesvio(ZZLugar, 8)) Then
                                                                XXCanti = Val(ZCargaDesvio(ZZLugar, 13))
                                                            End If
                                                        
                                                            If Xlugar < 10 And XCanti > 0 Then
                                                            
                                                                If XXCanti > 0 Then
                                                                
                                                                    Xlugar = Xlugar + 1
                                                                    If rstMovguia!Saldo > XXCanti Then
                                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                                        Impre(Xlugar, 2) = XXCanti
                                                                        XCanti = XCanti - XXCanti
                                                                            Else
                                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                                        Impre(Xlugar, 2) = rstMovguia!Saldo
                                                                        XCanti = XCanti - rstMovguia!Saldo
                                                                    End If
                                                                
                                                                        Else
                                                            
                                                                    Xlugar = Xlugar + 1
                                                                    If rstMovguia!Saldo > XCanti Then
                                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                                        Impre(Xlugar, 2) = XCanti
                                                                        XCanti = 0
                                                                            Else
                                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                                        Impre(Xlugar, 2) = rstMovguia!Saldo
                                                                        XCanti = XCanti - rstMovguia!Saldo
                                                                    End If
                                                                    
                                                                End If
                                                            End If
                                                        End If
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
                                rstMovguia.Close
                            End If
                        
                        End If
                        
                        
                        
                        
                        
                        
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
                        ZSql = ZSql + " Order by Laudo.Laudo"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                        
                            With rstLaudo
                                .MoveFirst
                                If .NoMatch = False Then
                                    Do
                                        If .EOF = True Then
                                            Exit Do
                                        End If
                           
                                        WMarcaVencida = IIf(IsNull(rstLaudo!MarcaVencida), "", rstLaudo!MarcaVencida)
                                        WMarcaEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                                        QSaldo = rstLaudo!Saldo
                                        Call Redondeo(QSaldo)
                                        
                                        If QSaldo <> 0 And Trim(WMarcaVencida) = "" And Trim(WMarcaEstado) <> "N" Then
                                        
                                            If rstLaudo!Articulo = Articulo Then
                                            
                                                WLaudo = rstLaudo!Laudo
                                                ZEntra = "S"
                                                
                                                If WLaudo >= 190000 And WLaudo <= 194999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 990000 And WLaudo <= 994999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 290000 And WLaudo <= 294999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 390000 And WLaudo <= 394999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 490000 And WLaudo <= 494999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 590000 And WLaudo <= 594999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 690000 And WLaudo <= 694999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 790000 And WLaudo <= 794999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 890000 And WLaudo <= 894999 Then
                                                    ZEntra = "N"
                                                End If
                                                
                                                If ZEntra = "S" Then
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstLaudo!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = rstLaudo!Saldo
                                                            XCanti = XCanti - rstLaudo!Saldo
                                                        End If
                                                    End If
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
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
                        ZSql = ZSql + " Order by Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                           
                            With rstMovguia
                           
                                .MoveFirst
                           
                                If .NoMatch = False Then
                                    Do
                           
                                        If .EOF = True Then
                                            Exit Do
                                        End If
                        
                                        WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                        
                                        QSaldo = rstMovguia!Saldo
                                        Call Redondeo(QSaldo)
                                        
                                        If QSaldo <> 0 And Trim(WMarcaVencida) = "" And Trim(WMarcaEstado) <> "N" Then
                                        
                                            If rstMovguia!Articulo = Articulo Then
                                            
                                                WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                                                ZEntra = "S"
                                                    
                                                If WLaudo >= 190000 And WLaudo <= 194999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 990000 And WLaudo <= 994999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 290000 And WLaudo <= 294999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 390000 And WLaudo <= 394999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 490000 And WLaudo <= 494999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 590000 And WLaudo <= 594999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 690000 And WLaudo <= 694999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 790000 And WLaudo <= 794999 Then
                                                    ZEntra = "N"
                                                End If
                                                If WLaudo >= 890000 And WLaudo <= 894999 Then
                                                    ZEntra = "N"
                                                End If
                                            
                                                If ZEntra = "S" Then
                                            
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
                                                    End If
                                                    
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
                            rstMovguia.Close
                        End If
                        
                    End If
                    
                    WArticulo1 = Left$(Articulo, 2)
                    WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
                    WCantidad = Cantidad
                    
                    For ZCiclo = 1 To 10
                        ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                        ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                    Next ZCiclo
                    
                    Linea = Linea + 1
                    WLinea = Str$(Linea)
       
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO ImpreHoja ("
                    ZSql = ZSql & "Hoja ,"
                    ZSql = ZSql & "Renglon ,"
                    ZSql = ZSql & "Fecha ,"
                    ZSql = ZSql & "Codigo1 ,"
                    ZSql = ZSql & "Codigo2 ,"
                    ZSql = ZSql & "Maquina ,"
                    ZSql = ZSql & "Articulo1 ,"
                    ZSql = ZSql & "Articulo2 ,"
                    ZSql = ZSql & "Cantidad ,"
                    ZSql = ZSql & "Canti1 ,"
                    ZSql = ZSql & "Lote1 ,"
                    ZSql = ZSql & "Canti2 ,"
                    ZSql = ZSql & "Lote2 ,"
                    ZSql = ZSql & "Canti3 ,"
                    ZSql = ZSql & "Lote3 ,"
                    ZSql = ZSql & "Teorico ,"
                    ZSql = ZSql & "Metodo ,"
                    ZSql = ZSql & "Efluentes ,"
                    ZSql = ZSql & "DesEfluentesI ,"
                    ZSql = ZSql & "DesEfluentesII ,"
                    ZSql = ZSql & "VersionI ,"
                    ZSql = ZSql & "VersionII ,"
                    ZSql = ZSql & "VersionIII ,"
                    ZSql = ZSql & "Equipo ,"
                    ZSql = ZSql & "Metodo1 ,"
                    ZSql = ZSql & "Metodo2 ,"
                    ZSql = ZSql & "Metodo3 ,"
                    ZSql = ZSql & "Metodo4 ,"
                    ZSql = ZSql & "Metodo5 ,"
                    ZSql = ZSql & "Metodo6 ,"
                    ZSql = ZSql & "Metodo7 ,"
                    ZSql = ZSql & "Metodo8 ,"
                    ZSql = ZSql & "Metodo9 ,"
                    ZSql = ZSql & "Metodo10 ,"
                    ZSql = ZSql & "Especificacion1 ,"
                    ZSql = ZSql & "Especificacion2 ,"
                    ZSql = ZSql & "Especificacion3 ,"
                    ZSql = ZSql & "Especificacion4 ,"
                    ZSql = ZSql & "Especificacion5 ,"
                    ZSql = ZSql & "Especificacion6 ,"
                    ZSql = ZSql & "Especificacion7 ,"
                    ZSql = ZSql & "Especificacion8 ,"
                    ZSql = ZSql & "Especificacion9 ,"
                    ZSql = ZSql & "Especificacion10 ,"
                    ZSql = ZSql & "DesMetodo1 ,"
                    ZSql = ZSql & "DesMetodo2 ,"
                    ZSql = ZSql & "DesMetodo3 ,"
                    ZSql = ZSql & "DesMetodo4 ,"
                    ZSql = ZSql & "DesMetodo5 ,"
                    ZSql = ZSql & "DesMetodo6 ,"
                    ZSql = ZSql & "DesMetodo7 ,"
                    ZSql = ZSql & "DesMetodo8 ,"
                    ZSql = ZSql & "DesMetodo9 ,"
                    ZSql = ZSql & "DesMetodo10 )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + WHoja + "',"
                    ZSql = ZSql & "'" + WLinea + "',"
                    ZSql = ZSql & "'" + WFecha + "',"
                    ZSql = ZSql & "'" + WCodigo1 + "',"
                    ZSql = ZSql & "'" + WCodigo2 + "',"
                    ZSql = ZSql & "'" + WMaquina + "',"
                    ZSql = ZSql & "'" + WArticulo1 + "',"
                    ZSql = ZSql & "'" + WArticulo2 + "',"
                    ZSql = ZSql & "'" + WCantidad + "',"
                    ZSql = ZSql & "'" + ZZCanti(1) + "',"
                    ZSql = ZSql & "'" + ZZLote(1) + "',"
                    ZSql = ZSql & "'" + ZZCanti(2) + "',"
                    ZSql = ZSql & "'" + ZZLote(2) + "',"
                    ZSql = ZSql & "'" + ZZCanti(3) + "',"
                    ZSql = ZSql & "'" + ZZLote(3) + "',"
                    ZSql = ZSql & "'" + WTeorico + "',"
                    ZSql = ZSql & "'" + ZMetodo + "',"
                    ZSql = ZSql & "'" + ZEfluentes + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                    ZSql = ZSql & "'" + ZVersionI + "',"
                    ZSql = ZSql & "'" + ZVersionII + "',"
                    ZSql = ZSql & "'" + ZVersionIII + "',"
                    ZSql = ZSql & "'" + Equipo.Text + "',"
                    ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
    
                    spImpreHoja = ZSql
                    Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                    For ZCiclo = 1 To 10
                    
                        If Val(ZZCanti(ZCiclo)) <> 0 Then
                    
                            LineaII = LineaII + 1
                            WLIneaII = Str$(LineaII)
                        
                            WCantidadII = ZZCanti(ZCiclo)
                            WLoteII = ZZLote(ZCiclo)
                            
                            ZSql = ""
                            ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                            ZSql = ZSql & "Hoja ,"
                            ZSql = ZSql & "Renglon ,"
                            ZSql = ZSql & "Fecha ,"
                            ZSql = ZSql & "Codigo1 ,"
                            ZSql = ZSql & "Codigo2 ,"
                            ZSql = ZSql & "Maquina ,"
                            ZSql = ZSql & "Articulo1 ,"
                            ZSql = ZSql & "Articulo2 ,"
                            ZSql = ZSql & "Cantidad ,"
                            ZSql = ZSql & "Lote ,"
                            ZSql = ZSql & "Observaciones ,"
                            ZSql = ZSql & "Teorico ,"
                            ZSql = ZSql & "Terminado ,"
                            ZSql = ZSql & "Equipo )"
                            ZSql = ZSql & "Values ("
                            ZSql = ZSql & "'" + WHoja + "',"
                            ZSql = ZSql & "'" + WLIneaII + "',"
                            ZSql = ZSql & "'" + WFecha + "',"
                            ZSql = ZSql & "'" + WCodigo1 + "',"
                            ZSql = ZSql & "'" + WCodigo2 + "',"
                            ZSql = ZSql & "'" + WMaquina + "',"
                            ZSql = ZSql & "'" + WArticulo1 + "',"
                            ZSql = ZSql & "'" + WArticulo2 + "',"
                            ZSql = ZSql & "'" + WCantidadII + "',"
                            ZSql = ZSql & "'" + WLoteII + "',"
                            ZSql = ZSql & "'" + ZZZObserva + "',"
                            ZSql = ZSql & "'" + WTeorico + "',"
                            ZSql = ZSql & "'" + Producto.Text + "',"
                            ZSql = ZSql & "'" + Equipo.Text + "')"
    
                            spImpreHojaII = ZSql
                            Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                        
                    Next ZCiclo
                    
                End If

                If Tipo = "T" Then
                
                    Erase Impre
                    Xlugar = 0
                    XCanti = Val(Cantidad)
                    
                    ZLugar = A
                    Impre(1, 1) = Val(ZLote(ZLugar, 1))
                    Impre(1, 2) = Val(ZLote(ZLugar, 2))
                    Impre(2, 1) = Val(ZLote(ZLugar, 3))
                    Impre(2, 2) = Val(ZLote(ZLugar, 4))
                    Impre(3, 1) = Val(ZLote(ZLugar, 5))
                    Impre(3, 2) = Val(ZLote(ZLugar, 6))
                    
                    If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                    
                        Impre(1, 1) = 0
                        Impre(1, 2) = 0
                        Impre(2, 1) = 0
                        Impre(2, 2) = 0
                        Impre(3, 1) = 0
                        Impre(3, 2) = 0
                        Impre(4, 1) = 0
                        Impre(4, 2) = 0
                        Impre(5, 1) = 0
                        Impre(5, 2) = 0
                        Impre(6, 1) = 0
                        Impre(6, 2) = 0
                        Impre(7, 1) = 0
                        Impre(7, 2) = 0
                        Impre(8, 1) = 0
                        Impre(8, 2) = 0
                        Impre(9, 1) = 0
                        Impre(9, 2) = 0
                        Impre(10, 1) = 0
                        Impre(10, 2) = 0
                    
                        XParam = "'" + Terminado + "','" _
                                     + Terminado + "'"
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
                        
                                        WMarcaEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                                        
                                        If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 And Trim(WMarcaEstado) <> "N" Then
                                            If rstHoja!Producto = Terminado Then
                                                If Xlugar < 10 And XCanti > 0 Then
                                                    Xlugar = Xlugar + 1
                                                    If rstHoja!Saldo > XCanti Then
                                                        Impre(Xlugar, 1) = rstHoja!Hoja
                                                        Impre(Xlugar, 2) = XCanti
                                                        XCanti = 0
                                                            Else
                                                        Impre(Xlugar, 1) = rstHoja!Hoja
                                                        Impre(Xlugar, 2) = rstHoja!Saldo
                                                        XCanti = XCanti - rstHoja!Saldo
                                                    End If
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
                            rstHoja.Close
                        End If
                    
                        XParam = "'" + Terminado + "','" _
                                     + Terminado + "'"
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
                        
                                        WMarcaEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                        
                                        If rstMovguia!Saldo <> 0 And Trim(WMarcaEstado) <> "N" Then
                                            If rstMovguia!Terminado = Terminado Then
                                                If Xlugar < 10 And XCanti > 0 Then
                                                    Xlugar = Xlugar + 1
                                                    If rstMovguia!Saldo > XCanti Then
                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                        Impre(Xlugar, 2) = XCanti
                                                        XCanti = 0
                                                            Else
                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                        Impre(Xlugar, 2) = rstMovguia!Saldo
                                                        XCanti = XCanti - rstMovguia!Saldo
                                                    End If
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
                            rstMovguia.Close
                        End If
                    
                    End If

                    Linea = Linea + 1
                    WLinea = Str$(Linea)
                    
                    WArticulo1 = Left$(Terminado, 2)
                    WArticulo2 = Mid$(Terminado, 4, 5) + "-" + Right$(Terminado, 3)
                    WCantidad = Cantidad
                    
                    For ZCiclo = 1 To 10
                        ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                        ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                    Next ZCiclo
                    
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO ImpreHoja ("
                    ZSql = ZSql & "Hoja ,"
                    ZSql = ZSql & "Renglon ,"
                    ZSql = ZSql & "Fecha ,"
                    ZSql = ZSql & "Codigo1 ,"
                    ZSql = ZSql & "Codigo2 ,"
                    ZSql = ZSql & "Maquina ,"
                    ZSql = ZSql & "Articulo1 ,"
                    ZSql = ZSql & "Articulo2 ,"
                    ZSql = ZSql & "Cantidad ,"
                    ZSql = ZSql & "Canti1 ,"
                    ZSql = ZSql & "Lote1 ,"
                    ZSql = ZSql & "Canti2 ,"
                    ZSql = ZSql & "Lote2 ,"
                    ZSql = ZSql & "Canti3 ,"
                    ZSql = ZSql & "Lote3 ,"
                    ZSql = ZSql & "Teorico ,"
                    ZSql = ZSql & "Metodo ,"
                    ZSql = ZSql & "Efluentes ,"
                    ZSql = ZSql & "DesEfluentesI ,"
                    ZSql = ZSql & "DesEfluentesII ,"
                    ZSql = ZSql & "VersionI ,"
                    ZSql = ZSql & "VersionII ,"
                    ZSql = ZSql & "VersionIII ,"
                    ZSql = ZSql & "Equipo ,"
                    ZSql = ZSql & "Metodo1 ,"
                    ZSql = ZSql & "Metodo2 ,"
                    ZSql = ZSql & "Metodo3 ,"
                    ZSql = ZSql & "Metodo4 ,"
                    ZSql = ZSql & "Metodo5 ,"
                    ZSql = ZSql & "Metodo6 ,"
                    ZSql = ZSql & "Metodo7 ,"
                    ZSql = ZSql & "Metodo8 ,"
                    ZSql = ZSql & "Metodo9 ,"
                    ZSql = ZSql & "Metodo10 ,"
                    ZSql = ZSql & "Especificacion1 ,"
                    ZSql = ZSql & "Especificacion2 ,"
                    ZSql = ZSql & "Especificacion3 ,"
                    ZSql = ZSql & "Especificacion4 ,"
                    ZSql = ZSql & "Especificacion5 ,"
                    ZSql = ZSql & "Especificacion6 ,"
                    ZSql = ZSql & "Especificacion7 ,"
                    ZSql = ZSql & "Especificacion8 ,"
                    ZSql = ZSql & "Especificacion9 ,"
                    ZSql = ZSql & "Especificacion10 ,"
                    ZSql = ZSql & "DesMetodo1 ,"
                    ZSql = ZSql & "DesMetodo2 ,"
                    ZSql = ZSql & "DesMetodo3 ,"
                    ZSql = ZSql & "DesMetodo4 ,"
                    ZSql = ZSql & "DesMetodo5 ,"
                    ZSql = ZSql & "DesMetodo6 ,"
                    ZSql = ZSql & "DesMetodo7 ,"
                    ZSql = ZSql & "DesMetodo8 ,"
                    ZSql = ZSql & "DesMetodo9 ,"
                    ZSql = ZSql & "DesMetodo10 )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + WHoja + "',"
                    ZSql = ZSql & "'" + WLinea + "',"
                    ZSql = ZSql & "'" + WFecha + "',"
                    ZSql = ZSql & "'" + WCodigo1 + "',"
                    ZSql = ZSql & "'" + WCodigo2 + "',"
                    ZSql = ZSql & "'" + WMaquina + "',"
                    ZSql = ZSql & "'" + WArticulo1 + "',"
                    ZSql = ZSql & "'" + WArticulo2 + "',"
                    ZSql = ZSql & "'" + WCantidad + "',"
                    ZSql = ZSql & "'" + ZZCanti(1) + "',"
                    ZSql = ZSql & "'" + ZZLote(1) + "',"
                    ZSql = ZSql & "'" + ZZCanti(2) + "',"
                    ZSql = ZSql & "'" + ZZLote(2) + "',"
                    ZSql = ZSql & "'" + ZZCanti(3) + "',"
                    ZSql = ZSql & "'" + ZZLote(3) + "',"
                    ZSql = ZSql & "'" + WTeorico + "',"
                    ZSql = ZSql & "'" + ZMetodo + "',"
                    ZSql = ZSql & "'" + ZEfluentes + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                    ZSql = ZSql & "'" + ZVersionI + "',"
                    ZSql = ZSql & "'" + ZVersionII + "',"
                    ZSql = ZSql & "'" + ZVersionIII + "',"
                    ZSql = ZSql & "'" + Equipo.Text + "',"
                    ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
    
                    spImpreHoja = ZSql
                    Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                    Rem el problema esta aca
                    
                    
                    For ZCiclo = 1 To 10
                    
                        If Val(ZZCanti(ZCiclo)) <> 0 Then
                    
                            LineaII = LineaII + 1
                            WLIneaII = Str$(LineaII)
                        
                            WCantidadII = ZZCanti(ZCiclo)
                            WLoteII = ZZLote(ZCiclo)
                    
                            ZSql = ""
                            ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                            ZSql = ZSql & "Hoja ,"
                            ZSql = ZSql & "Renglon ,"
                            ZSql = ZSql & "Fecha ,"
                            ZSql = ZSql & "Codigo1 ,"
                            ZSql = ZSql & "Codigo2 ,"
                            ZSql = ZSql & "Maquina ,"
                            ZSql = ZSql & "Articulo1 ,"
                            ZSql = ZSql & "Articulo2 ,"
                            ZSql = ZSql & "Cantidad ,"
                            ZSql = ZSql & "Lote ,"
                            ZSql = ZSql & "Observaciones ,"
                            ZSql = ZSql & "Teorico ,"
                            ZSql = ZSql & "Terminado ,"
                            ZSql = ZSql & "Equipo )"
                            ZSql = ZSql & "Values ("
                            ZSql = ZSql & "'" + WHoja + "',"
                            ZSql = ZSql & "'" + WLIneaII + "',"
                            ZSql = ZSql & "'" + WFecha + "',"
                            ZSql = ZSql & "'" + WCodigo1 + "',"
                            ZSql = ZSql & "'" + WCodigo2 + "',"
                            ZSql = ZSql & "'" + WMaquina + "',"
                            ZSql = ZSql & "'" + WArticulo1 + "',"
                            ZSql = ZSql & "'" + WArticulo2 + "',"
                            ZSql = ZSql & "'" + WCantidadII + "',"
                            ZSql = ZSql & "'" + WLoteII + "',"
                            ZSql = ZSql & "'" + ZZZObserva + "',"
                            ZSql = ZSql & "'" + WTeorico + "',"
                            ZSql = ZSql & "'" + Producto.Text + "',"
                            ZSql = ZSql & "'" + Equipo.Text + "')"
        
                            spImpreHojaII = ZSql
                            Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                        
                    Next ZCiclo
                    
                End If
            
            Next A
            
            XLinea = Linea
            For Ciclo = XLinea To 14
            
                Linea = Linea + 1
                WLinea = Str$(Linea)
                        
                WArticulo1 = ""
                WArticulo2 = ""
                WCantidad = ""
                WCanti1 = ""
                WLote1 = ""
                WCanti2 = ""
                WLote2 = ""
                WCanti3 = ""
                WLote3 = ""
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ImpreHoja ("
                ZSql = ZSql & "Hoja ,"
                ZSql = ZSql & "Renglon ,"
                ZSql = ZSql & "Fecha ,"
                ZSql = ZSql & "Codigo1 ,"
                ZSql = ZSql & "Codigo2 ,"
                ZSql = ZSql & "Maquina ,"
                ZSql = ZSql & "Articulo1 ,"
                ZSql = ZSql & "Articulo2 ,"
                ZSql = ZSql & "Cantidad ,"
                ZSql = ZSql & "Canti1 ,"
                ZSql = ZSql & "Lote1 ,"
                ZSql = ZSql & "Canti2 ,"
                ZSql = ZSql & "Lote2 ,"
                ZSql = ZSql & "Canti3 ,"
                ZSql = ZSql & "Lote3 ,"
                ZSql = ZSql & "Teorico ,"
                ZSql = ZSql & "Metodo ,"
                ZSql = ZSql & "Efluentes ,"
                ZSql = ZSql & "DesEfluentesI ,"
                ZSql = ZSql & "DesEfluentesII ,"
                ZSql = ZSql & "VersionI ,"
                ZSql = ZSql & "VersionII ,"
                ZSql = ZSql & "VersionIII ,"
                ZSql = ZSql & "Equipo ,"
                ZSql = ZSql & "Metodo1 ,"
                ZSql = ZSql & "Metodo2 ,"
                ZSql = ZSql & "Metodo3 ,"
                ZSql = ZSql & "Metodo4 ,"
                ZSql = ZSql & "Metodo5 ,"
                ZSql = ZSql & "Metodo6 ,"
                ZSql = ZSql & "Metodo7 ,"
                ZSql = ZSql & "Metodo8 ,"
                ZSql = ZSql & "Metodo9 ,"
                ZSql = ZSql & "Metodo10 ,"
                ZSql = ZSql & "Especificacion1 ,"
                ZSql = ZSql & "Especificacion2 ,"
                ZSql = ZSql & "Especificacion3 ,"
                ZSql = ZSql & "Especificacion4 ,"
                ZSql = ZSql & "Especificacion5 ,"
                ZSql = ZSql & "Especificacion6 ,"
                ZSql = ZSql & "Especificacion7 ,"
                ZSql = ZSql & "Especificacion8 ,"
                ZSql = ZSql & "Especificacion9 ,"
                ZSql = ZSql & "Especificacion10 ,"
                ZSql = ZSql & "DesMetodo1 ,"
                ZSql = ZSql & "DesMetodo2 ,"
                ZSql = ZSql & "DesMetodo3 ,"
                ZSql = ZSql & "DesMetodo4 ,"
                ZSql = ZSql & "DesMetodo5 ,"
                ZSql = ZSql & "DesMetodo6 ,"
                ZSql = ZSql & "DesMetodo7 ,"
                ZSql = ZSql & "DesMetodo8 ,"
                ZSql = ZSql & "DesMetodo9 ,"
                ZSql = ZSql & "DesMetodo10 )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WHoja + "',"
                ZSql = ZSql & "'" + WLinea + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WCodigo1 + "',"
                ZSql = ZSql & "'" + WCodigo2 + "',"
                ZSql = ZSql & "'" + WMaquina + "',"
                ZSql = ZSql & "'" + WArticulo1 + "',"
                ZSql = ZSql & "'" + WArticulo2 + "',"
                ZSql = ZSql & "'" + WCantidad + "',"
                ZSql = ZSql & "'" + ZCanti1 + "',"
                ZSql = ZSql & "'" + ZLote1 + "',"
                ZSql = ZSql & "'" + ZCanti2 + "',"
                ZSql = ZSql & "'" + ZLote2 + "',"
                ZSql = ZSql & "'" + ZCanti3 + "',"
                ZSql = ZSql & "'" + ZLote3 + "',"
                ZSql = ZSql & "'" + WTeorico + "',"
                ZSql = ZSql & "'" + ZMetodo + "',"
                ZSql = ZSql & "'" + ZEfluentes + "',"
                ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                ZSql = ZSql & "'" + ZVersionI + "',"
                ZSql = ZSql & "'" + ZVersionII + "',"
                ZSql = ZSql & "'" + ZVersionIII + "',"
                ZSql = ZSql & "'" + Equipo.Text + "',"
                ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
        
                spImpreHoja = ZSql
                Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            
            XLinea = LineaII
            For Ciclo = XLinea To 24
            
                LineaII = LineaII + 1
                WLIneaII = Str$(LineaII)
                        
                WCantidadII = ""
                WLoteII = ""
                                                   
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                ZSql = ZSql & "Hoja ,"
                ZSql = ZSql & "Renglon ,"
                ZSql = ZSql & "Fecha ,"
                ZSql = ZSql & "Codigo1 ,"
                ZSql = ZSql & "Codigo2 ,"
                ZSql = ZSql & "Maquina ,"
                ZSql = ZSql & "Articulo1 ,"
                ZSql = ZSql & "Articulo2 ,"
                ZSql = ZSql & "Cantidad ,"
                ZSql = ZSql & "Lote ,"
                ZSql = ZSql & "Observaciones ,"
                ZSql = ZSql & "Teorico ,"
                ZSql = ZSql & "Terminado ,"
                ZSql = ZSql & "Equipo )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WHoja + "',"
                ZSql = ZSql & "'" + WLIneaII + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WCodigo1 + "',"
                ZSql = ZSql & "'" + WCodigo2 + "',"
                ZSql = ZSql & "'" + WMaquina + "',"
                ZSql = ZSql & "'" + WArticulo1 + "',"
                ZSql = ZSql & "'" + WArticulo2 + "',"
                ZSql = ZSql & "'" + WCantidadII + "',"
                ZSql = ZSql & "'" + WLoteII + "',"
                ZSql = ZSql & "'" + "" + "',"
                ZSql = ZSql & "'" + WTeorico + "',"
                ZSql = ZSql & "'" + Producto.Text + "',"
                ZSql = ZSql & "'" + Equipo.Text + "')"
        
                spImpreHojaII = ZSql
                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE ImpreHoja SET "
            ZSql = ZSql + " Especificacion11 = " + "'" + ZZEspecificacion(11) + "',"
            ZSql = ZSql + " Especificacion22 = " + "'" + ZZEspecificacion(12) + "',"
            ZSql = ZSql + " Especificacion33 = " + "'" + ZZEspecificacion(13) + "',"
            ZSql = ZSql + " Especificacion44 = " + "'" + ZZEspecificacion(14) + "',"
            ZSql = ZSql + " Especificacion55 = " + "'" + ZZEspecificacion(15) + "',"
            ZSql = ZSql + " Especificacion66 = " + "'" + ZZEspecificacion(16) + "',"
            ZSql = ZSql + " Especificacion77 = " + "'" + ZZEspecificacion(17) + "',"
            ZSql = ZSql + " Especificacion88 = " + "'" + ZZEspecificacion(18) + "',"
            ZSql = ZSql + " Especificacion99 = " + "'" + ZZEspecificacion(19) + "',"
            ZSql = ZSql + " Especificacion1010 = " + "'" + ZZEspecificacion(20) + "'"
            spImpreHoja = ZSql
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Listado.WindowTitle = "Impresion de Hoja de Produccion"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
   
            Listado.Destination = 1
            Rem Listado.Destination = 0
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Select Case XTipoPro
                Case "CO"
                    Listado.ReportFileName = "ImpreHojaNuevoA4.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} in 0 to 10 and {ImpreHoja.Hoja} in 0 to 999999"
                    
                    Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10 " _
                            + "From " _
                            + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                            + "Where " _
                            + "ImpreHoja.Hoja >= 0 AND " _
                            + "ImpreHoja.Hoja <= 999999 AND " _
                            + "ImpreHoja.Renglon >= 0 AND " _
                            + "ImpreHoja.Renglon <= 10"
            
                Case Else
                    Listado.ReportFileName = "ImpreHojaNuevo.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
                
                    Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10 " _
                            + "From " _
                            + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                            + "Where " _
                            + "ImpreHoja.Hoja >= 0 AND " _
                            + "ImpreHoja.Hoja <= 999999"
            End Select
    
            Listado.Connect = Connect()
            Listado.Action = 1
            
            T$ = "Hoja de Produccion"
            m$ = "Desea Imprimir la Hoja del Almacenero"
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Select Case XTipoPro
                    Case "CO"
                        Listado.ReportFileName = "ImpreHojaNuevoIIA4.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                                + "From " _
                                + DSQ + ".dbo.ImpreHojaII ImpreHojaII " _
                                + "Where " _
                                + "ImpreHojaII.Hoja >= 0 AND " _
                                + "ImpreHojaII.Hoja <= 999999"
            
                    Case Else
                        Listado.ReportFileName = "ImpreHojaNuevoII.rpt"
                        Rem Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
                        Listado.GroupSelectionFormula = ""
                        Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Renglon, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo , ImpreHojaII.Observaciones " _
                        + "From " _
                        + DSQ + ".dbo.ImpreHojaII ImprehojaII " _
                        + "Where ImpreHojaII.Hoja >= 0 AND ImpreHojaII.Hoja <= 999999"
                        
                        
                        Rem SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.VersionI, ImpreHoja.VersionII, ImpreHoja.VersionIII, ImpreHoja.Equipo, ImpreHoja.Metodo1, ImpreHoja.Metodo2, ImpreHoja.Metodo3, ImpreHoja.Metodo4, ImpreHoja.Metodo5, ImpreHoja.Metodo6, ImpreHoja.Metodo7, ImpreHoja.Metodo8, ImpreHoja.Metodo9, ImpreHoja.Metodo10, ImpreHoja.Especificacion1, ImpreHoja.Especificacion2, ImpreHoja.Especificacion3, ImpreHoja.Especificacion4, ImpreHoja.Especificacion5, ImpreHoja.Especificacion6, ImpreHoja.Especificacion7, ImpreHoja.Especificacion8, ImpreHoja.Especificacion9, ImpreHoja.Especificacion10, " _
                            REM        + "ImpreHoja.Especificacion11, ImpreHoja.Especificacion22, ImpreHoja.Especificacion33, ImpreHoja.Especificacion44, ImpreHoja.Especificacion55, ImpreHoja.Especificacion66, ImpreHoja.Especificacion77, ImpreHoja.Especificacion88, ImpreHoja.Especificacion99, ImpreHoja.Especificacion1010 " _
                               REM     + "From " _
                                  REM  + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    REM+ "Where " _
                                    REM+ "ImpreHoja.Hoja >= 0 AND " _
                                    REM+ "ImpreHoja.Hoja <= 999999"
                End Select
                Listado.Connect = Connect()
                Listado.Action = 1
            End If
            
            If Trim(Equipo.Text) <> "" Then
                Select Case XTipoPro
                    Case "CO"
                        Listado.ReportFileName = "ImpreHojaNuevoIIIA4.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} = 1 and {ImpreHoja.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.DesEfluentesII, ImpreHoja.Equipo " _
                                    + "From " _
                                    + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    + "Where " _
                                    + "ImpreHoja.Hoja >= 0 AND " _
                                    + "ImpreHoja.Hoja<= 999999 AND " _
                                    + "ImpreHoja.Renglon = 1"
                        Listado.Connect = Connect()
                        Listado.Action = 1
            
                    Case Else
                        Listado.ReportFileName = "ImpreHojaNuevoIII.rpt"
                        Listado.GroupSelectionFormula = "{ImpreHoja.Renglon} = 1 and {ImpreHoja.Hoja} in 0 to 999999"
                        
                        Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Renglon, ImpreHoja.Fecha, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Teorico, ImpreHoja.Metodo, ImpreHoja.DesEfluentesI, ImpreHoja.DesEfluentesII, ImpreHoja.Equipo " _
                                    + "From " _
                                    + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                                    + "Where " _
                                    + "ImpreHoja.Hoja >= 0 AND " _
                                    + "ImpreHoja.Hoja <= 999999 AND " _
                                    + "ImpreHoja.Renglon = 1"
                        Listado.Connect = Connect()
                        Listado.Action = 1
                End Select
            End If
        
        
        
        Case 5
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WEscrito = IIf(IsNull(rstTerminado!Escrito), "0", rstTerminado!Escrito)
                rstTerminado.Close
            End If
            
            If WEscrito <> 0 Then
            
                If WEscrito = 1 Then
                    Call Impresion_Farma
                        Else
                    Call Impresion_FarmaII
                End If
                
                T$ = "Hoja de Produccion"
                m$ = "Desea Imprimir la Hoja del Almacenero"
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    Call ImpreHojaAlmacero
                    Call ImpreHojaDesvio
                End If
            
                    Else
                    
             Rem BY NAN 24-11 ES PARA LPT1
                Open "lpt1" For Output As #1

                Print #1, Chr$(27) + Chr$(71);
                Print #1, Chr$(18)

                Print #1, Tab(22); Left$(Producto.Text, 2);
                Print #1, Tab(77); "SIII"

                Print #1, Tab(8); Fecha.Text;
                Print #1, Tab(19); Alinea("#####", Mid$(Producto.Text, 4, 5));
                Print #1, "/"; Right$(Producto.Text, 3);
                Print #1, Tab(33); Chr$(14); Alinea("######", Hoja.Text)

                Print #1,
                Print #1,
                Print #1,

                Linea = 0
        
                For A = 1 To 40
        
                    Tipo = UCase(Grilla.TextMatrix(A, 1))
                    Terminado = UCase(Grilla.TextMatrix(A, 2))
                    Articulo = UCase(Grilla.TextMatrix(A, 3))
                    Cantidad = Grilla.TextMatrix(A, 5)
                 
                    If Tipo = "M" Then
                
                        Rem PROCESA LOS LAUDOS

                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                    
                        ZLugar = A
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                    
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                    
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0

                            XParam = "'" + Articulo + "','" _
                                     + Articulo + "'"
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
                               
                                            If rstLaudo!Saldo <> 0 Then
                                                If rstLaudo!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstLaudo!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = rstLaudo!Saldo
                                                            XCanti = XCanti - rstLaudo!Saldo
                                                        End If
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
                    
                            XParam = "'" + Articulo + "','" _
                                     + Articulo + "'"
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
                               
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Articulo = Articulo Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                                rstMovguia.Close
                            End If
                        
                        End If
                    
                        Linea = Linea + 1

                        Print #1, Tab(13); Left$(Articulo, 2);
                        Print #1, Tab(18); Mid$(Articulo, 4, 3);
                        Print #1, "-";
                        Print #1, Right$(Articulo, 3);
                        If Val(Teorico.Text) < 100 Then
                            Print #1, Tab(27); Alinea("###.##", Cantidad);
                                Else
                            Print #1, Tab(27); Alinea("####.#", Cantidad);
                        End If
                    
                        If Impre(1, 2) <> 0 Then
                            If Impre(1, 2) < 100 Then
                                Print #1, Tab(34); Alinea("###.##", Str$(Impre(1, 2)));
                                    Else
                                Print #1, Tab(34); Alinea("####.#", Str$(Impre(1, 2)));
                            End If
                        End If
                        If Impre(1, 1) <> 0 Then
                            Print #1, Tab(41); Alinea("######", Str$(Impre(1, 1)));
                        End If
                    
                        If Impre(2, 2) <> 0 Then
                            If Impre(2, 2) < 100 Then
                                Print #1, Tab(48); Alinea("###.##", Str$(Impre(2, 2)));
                                    Else
                                Print #1, Tab(48); Alinea("####.#", Str$(Impre(2, 2)));
                            End If
                        End If
                        If Impre(2, 1) <> 0 Then
                            Print #1, Tab(55); Alinea("######", Str$(Impre(2, 1)));
                        End If
                    
                        If Impre(3, 2) <> 0 Then
                            If Impre(3, 2) < 100 Then
                                Print #1, Tab(62); Alinea("###.##", Str$(Impre(3, 2)));
                                    Else
                                Print #1, Tab(62); Alinea("####.#", Str$(Impre(3, 2)));
                            End If
                        End If
                        If Impre(3, 1) <> 0 Then
                            Print #1, Tab(69); Alinea("######", Str$(Impre(3, 1)));
                        End If
                        
                        Print #1,
                        Print #1,

                    End If

                    If Tipo = "T" Then
                
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                    
                        ZLugar = A
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                    
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                    
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                    
                            XParam = "'" + Terminado + "','" _
                                     + Terminado + "'"
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
                        
                                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                                If rstHoja!Producto = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstHoja!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                                            XCanti = XCanti - rstHoja!Saldo
                                                        End If
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
                                rstHoja.Close
                            End If
                    
                            XParam = "'" + Terminado + "','" _
                                     + Terminado + "'"
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
                        
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Terminado = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                                rstMovguia.Close
                            End If
                    
                        End If

                        Linea = Linea + 1

                        Print #1, Tab(13); Left$(Terminado, 2);
                        Print #1, Tab(18); Mid$(Terminado, 4, 5);
                        Print #1, "-";
                        Print #1, Right$(Terminado, 3);
                        If Val(Teorico.Text) < 100 Then
                            Print #1, Tab(27); Alinea("###.##", Cantidad);
                                Else
                            Print #1, Tab(27); Alinea("####.#", Cantidad);
                        End If
                    
                        If Impre(1, 2) <> 0 Then
                            If Impre(1, 2) < 100 Then
                                Print #1, Tab(34); Alinea("###.##", Str$(Impre(1, 2)));
                                    Else
                                Print #1, Tab(34); Alinea("####.#", Str$(Impre(1, 2)));
                            End If
                        End If
                        If Impre(1, 1) <> 0 Then
                            Print #1, Tab(41); Alinea("######", Str$(Impre(1, 1)));
                        End If
                    
                        If Impre(2, 2) <> 0 Then
                            If Impre(2, 2) < 100 Then
                                Print #1, Tab(47); Alinea("###.##", Str$(Impre(2, 2)));
                                    Else
                                Print #1, Tab(47); Alinea("####.#", Str$(Impre(2, 2)));
                            End If
                        End If
                        If Impre(2, 1) <> 0 Then
                            Print #1, Tab(53); Alinea("######", Str$(Impre(2, 1)));
                        End If
                    
                        If Impre(3, 2) <> 0 Then
                            If Impre(3, 2) < 100 Then
                                Print #1, Tab(61); Alinea("###.##", Str$(Impre(3, 2)));
                                    Else
                                Print #1, Tab(61); Alinea("####.#", Str$(Impre(3, 2)));
                            End If
                        End If
                        If Impre(3, 1) <> 0 Then
                            Print #1, Tab(69); Alinea("######", Str$(Impre(3, 1)));
                        End If
                    
                        Print #1,
                        Print #1,

                    End If
            
                Next A

                For Ciclo = Linea To 14

                    Print #1,
                    Print #1,

                Next Ciclo

                Print #1, Tab(27); Alinea("####.#", Teorico.Text)



                Print #1,
               Print #1, Chr$(27) + Chr$(72)
                Print #1, Chr$(12)
        
                Close #1
                
            End If
        
        Case Else

            Sql1 = "DELETE ImpreHoja"
            spImpreHoja = Sql1
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
            Linea = 0
            ZLinea = 0
        
            For A = 1 To 40
        
                Tipo = UCase(Grilla.TextMatrix(A, 1))
                Terminado = UCase(Grilla.TextMatrix(A, 2))
                Articulo = UCase(Grilla.TextMatrix(A, 3))
                Cantidad = Grilla.TextMatrix(A, 5)
                WHoja = Hoja.Text
                WArticulo = Right(Articulo, 7)
                WTerminado = Terminado
                WCantidad = Cantidad
                Linea = Linea + 1
                WLinea = Str$(Linea)
                WTeorico = Teorico.Text
                
                If Tipo = "M" Then
                
                    Rem PROCESA LOS LAUDOS

                    Erase Impre
                    Xlugar = 0
                    XCanti = Val(Cantidad)
                    
                    ZLugar = A
                    Impre(1, 1) = Val(ZLote(ZLugar, 1))
                    Impre(1, 2) = Val(ZLote(ZLugar, 2))
                    Impre(2, 1) = Val(ZLote(ZLugar, 3))
                    Impre(2, 2) = Val(ZLote(ZLugar, 4))
                    Impre(3, 1) = Val(ZLote(ZLugar, 5))
                    Impre(3, 2) = Val(ZLote(ZLugar, 6))
                    
                    If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                    
                        Impre(1, 1) = 0
                        Impre(1, 2) = 0
                        Impre(2, 1) = 0
                        Impre(2, 2) = 0
                        Impre(3, 1) = 0
                        Impre(3, 2) = 0

                        XParam = "'" + Articulo + "','" _
                                     + Articulo + "'"
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
                           
                                        If rstLaudo!Saldo <> 0 Then
                                            If rstLaudo!Articulo = Articulo Then
                                                If Xlugar < 3 And XCanti > 0 Then
                                                    Xlugar = Xlugar + 1
                                                    If rstLaudo!Saldo > XCanti Then
                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                        Impre(Xlugar, 2) = XCanti
                                                        XCanti = 0
                                                            Else
                                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                                        Impre(Xlugar, 2) = rstLaudo!Saldo
                                                        XCanti = XCanti - rstLaudo!Saldo
                                                    End If
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
                    
                        XParam = "'" + Articulo + "','" _
                                     + Articulo + "'"
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
                           
                                        If rstMovguia!Saldo <> 0 Then
                                            If rstMovguia!Articulo = Articulo Then
                                                If Xlugar < 3 And XCanti > 0 Then
                                                    Xlugar = Xlugar + 1
                                                    If rstMovguia!Saldo > XCanti Then
                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                        Impre(Xlugar, 2) = XCanti
                                                        XCanti = 0
                                                            Else
                                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                                        Impre(Xlugar, 2) = rstMovguia!Saldo
                                                        XCanti = XCanti - rstMovguia!Saldo
                                                    End If
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
                            rstMovguia.Close
                        End If
                        
                    End If
                    
                    WCodigo1 = Left$(Producto.Text, 2)
                    WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
                    WArticulo1 = Left$(Articulo, 2)
                    WArticulo2 = Right$(Articulo, 7)
                    WFecha = Fecha.Text
                    WHoja = Hoja.Text
                    WTipo = Tipo
                    WLote1 = Impre(1, 1)
                    WCanti1 = Str$(Impre(1, 2))
                    WLote2 = Impre(2, 1)
                    WCanti2 = Str$(Impre(2, 2))
                    WLote3 = Impre(3, 1)
                    WCanti3 = Str$(Impre(3, 2))
              
                    
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO ImpreHoja ("
                    ZSql = ZSql & "Hoja ,"
                    ZSql = ZSql & "Renglon ,"
                    ZSql = ZSql & "Fecha ,"
                    ZSql = ZSql & "Codigo1 ,"
                    ZSql = ZSql & "Codigo2 ,"
                    ZSql = ZSql & "Maquina ,"
                    ZSql = ZSql & "Articulo1 ,"
                    ZSql = ZSql & "Articulo2 ,"
                    ZSql = ZSql & "Cantidad ,"
                    ZSql = ZSql & "Canti1 ,"
                    ZSql = ZSql & "Lote1 ,"
                    ZSql = ZSql & "Canti2 ,"
                    ZSql = ZSql & "Lote2 ,"
                    ZSql = ZSql & "Canti3 ,"
                    ZSql = ZSql & "Lote3 ,"
                    ZSql = ZSql & "Teorico ,"
                    ZSql = ZSql & "Metodo ,"
                    ZSql = ZSql & "Efluentes ,"
                    ZSql = ZSql & "DesEfluentesI ,"
                    ZSql = ZSql & "DesEfluentesII ,"
                    ZSql = ZSql & "VersionI ,"
                    ZSql = ZSql & "VersionII ,"
                    ZSql = ZSql & "VersionIII ,"
                    ZSql = ZSql & "Equipo ,"
                    ZSql = ZSql & "Metodo1 ,"
                    ZSql = ZSql & "Metodo2 ,"
                    ZSql = ZSql & "Metodo3 ,"
                    ZSql = ZSql & "Metodo4 ,"
                    ZSql = ZSql & "Metodo5 ,"
                    ZSql = ZSql & "Metodo6 ,"
                    ZSql = ZSql & "Metodo7 ,"
                    ZSql = ZSql & "Metodo8 ,"
                    ZSql = ZSql & "Metodo9 ,"
                    ZSql = ZSql & "Metodo10 ,"
                    ZSql = ZSql & "Especificacion1 ,"
                    ZSql = ZSql & "Especificacion2 ,"
                    ZSql = ZSql & "Especificacion3 ,"
                    ZSql = ZSql & "Especificacion4 ,"
                    ZSql = ZSql & "Especificacion5 ,"
                    ZSql = ZSql & "Especificacion6 ,"
                    ZSql = ZSql & "Especificacion7 ,"
                    ZSql = ZSql & "Especificacion8 ,"
                    ZSql = ZSql & "Especificacion9 ,"
                    ZSql = ZSql & "Especificacion10 ,"
                    ZSql = ZSql & "DesMetodo1 ,"
                    ZSql = ZSql & "DesMetodo2 ,"
                    ZSql = ZSql & "DesMetodo3 ,"
                    ZSql = ZSql & "DesMetodo4 ,"
                    ZSql = ZSql & "DesMetodo5 ,"
                    ZSql = ZSql & "DesMetodo6 ,"
                    ZSql = ZSql & "DesMetodo7 ,"
                    ZSql = ZSql & "DesMetodo8 ,"
                    ZSql = ZSql & "DesMetodo9 ,"
                    ZSql = ZSql & "DesMetodo10 )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + WHoja + "',"
                    ZSql = ZSql & "'" + WLinea + "',"
                    ZSql = ZSql & "'" + WFecha + "',"
                    ZSql = ZSql & "'" + WCodigo1 + "',"
                    ZSql = ZSql & "'" + WCodigo2 + "',"
                    ZSql = ZSql & "'" + WMaquina + "',"
                    ZSql = ZSql & "'" + WArticulo1 + "',"
                    ZSql = ZSql & "'" + WArticulo2 + "',"
                    ZSql = ZSql & "'" + WCantidad + "',"
                    ZSql = ZSql & "'" + WCanti1 + "',"
                    ZSql = ZSql & "'" + WLote1 + "',"
                    ZSql = ZSql & "'" + WCanti2 + "',"
                    ZSql = ZSql & "'" + WLote2 + "',"
                    ZSql = ZSql & "'" + WCanti3 + "',"
                    ZSql = ZSql & "'" + WLote3 + "',"
                    ZSql = ZSql & "'" + WTeorico + "',"
                    ZSql = ZSql & "'" + ZMetodo + "',"
                    ZSql = ZSql & "'" + ZEfluentes + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                    ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                    ZSql = ZSql & "'" + ZVersionI + "',"
                    ZSql = ZSql & "'" + ZVersionII + "',"
                    ZSql = ZSql & "'" + ZVersionIII + "',"
                    ZSql = ZSql & "'" + Equipo.Text + "',"
                    ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                    ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                    ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                    ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
                    spImpreHoja = ZSql
                    Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                    ZLinea = ZLinea + 1
                
                        Else
                
                    If Tipo = "T" Then
                    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = A
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
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
                            
                                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                                If rstHoja!Producto = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstHoja!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                                            XCanti = XCanti - rstHoja!Saldo
                                                        End If
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
                                rstHoja.Close
                            End If
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
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
                            
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Terminado = Terminado Then
                                                    If Xlugar < 3 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                                rstMovguia.Close
                            End If
                        
                        End If
                        
                        Rem WCodigo1 = Tipo
                        WArticulo1 = Left$(Terminado, 2)
                        WArticulo2 = Right$(Terminado, 7)
                        WFecha = Fecha.Text
                        WHoja = Hoja.Text
                        WTipo = Tipo
                        WLote1 = Impre(1, 1)
                        WCanti1 = Str$(Impre(1, 2))
                        WLote2 = Impre(2, 1)
                        WCanti2 = Str$(Impre(2, 2))
                        WLote3 = Impre(3, 1)
                        WCanti3 = Str$(Impre(3, 2))
                    
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO ImpreHoja ("
                        ZSql = ZSql & "Hoja ,"
                        ZSql = ZSql & "Renglon ,"
                        ZSql = ZSql & "Fecha ,"
                        ZSql = ZSql & "Codigo1 ,"
                        ZSql = ZSql & "Codigo2 ,"
                        ZSql = ZSql & "Maquina ,"
                        ZSql = ZSql & "Articulo1 ,"
                        ZSql = ZSql & "Articulo2 ,"
                        ZSql = ZSql & "Cantidad ,"
                        ZSql = ZSql & "Canti1 ,"
                        ZSql = ZSql & "Lote1 ,"
                        ZSql = ZSql & "Canti2 ,"
                        ZSql = ZSql & "Lote2 ,"
                        ZSql = ZSql & "Canti3 ,"
                        ZSql = ZSql & "Lote3 ,"
                        ZSql = ZSql & "Teorico ,"
                        ZSql = ZSql & "Metodo ,"
                        ZSql = ZSql & "Efluentes ,"
                        ZSql = ZSql & "DesEfluentesI ,"
                        ZSql = ZSql & "DesEfluentesII ,"
                        ZSql = ZSql & "VersionI ,"
                        ZSql = ZSql & "VersionII ,"
                        ZSql = ZSql & "VersionIII ,"
                        ZSql = ZSql & "Equipo ,"
                        ZSql = ZSql & "Metodo1 ,"
                        ZSql = ZSql & "Metodo2 ,"
                        ZSql = ZSql & "Metodo3 ,"
                        ZSql = ZSql & "Metodo4 ,"
                        ZSql = ZSql & "Metodo5 ,"
                        ZSql = ZSql & "Metodo6 ,"
                        ZSql = ZSql & "Metodo7 ,"
                        ZSql = ZSql & "Metodo8 ,"
                        ZSql = ZSql & "Metodo9 ,"
                        ZSql = ZSql & "Metodo10 ,"
                        ZSql = ZSql & "Especificacion1 ,"
                        ZSql = ZSql & "Especificacion2 ,"
                        ZSql = ZSql & "Especificacion3 ,"
                        ZSql = ZSql & "Especificacion4 ,"
                        ZSql = ZSql & "Especificacion5 ,"
                        ZSql = ZSql & "Especificacion6 ,"
                        ZSql = ZSql & "Especificacion7 ,"
                        ZSql = ZSql & "Especificacion8 ,"
                        ZSql = ZSql & "Especificacion9 ,"
                        ZSql = ZSql & "Especificacion10 ,"
                        ZSql = ZSql & "DesMetodo1 ,"
                        ZSql = ZSql & "DesMetodo2 ,"
                        ZSql = ZSql & "DesMetodo3 ,"
                        ZSql = ZSql & "DesMetodo4 ,"
                        ZSql = ZSql & "DesMetodo5 ,"
                        ZSql = ZSql & "DesMetodo6 ,"
                        ZSql = ZSql & "DesMetodo7 ,"
                        ZSql = ZSql & "DesMetodo8 ,"
                        ZSql = ZSql & "DesMetodo9 ,"
                        ZSql = ZSql & "DesMetodo10 )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WHoja + "',"
                        ZSql = ZSql & "'" + WLinea + "',"
                        ZSql = ZSql & "'" + WFecha + "',"
                        ZSql = ZSql & "'" + WCodigo1 + "',"
                        ZSql = ZSql & "'" + WCodigo2 + "',"
                        ZSql = ZSql & "'" + WMaquina + "',"
                        ZSql = ZSql & "'" + WArticulo1 + "',"
                        ZSql = ZSql & "'" + WArticulo2 + "',"
                        ZSql = ZSql & "'" + WCantidad + "',"
                        ZSql = ZSql & "'" + WCanti1 + "',"
                        ZSql = ZSql & "'" + WLote1 + "',"
                        ZSql = ZSql & "'" + WCanti2 + "',"
                        ZSql = ZSql & "'" + WLote2 + "',"
                        ZSql = ZSql & "'" + WCanti3 + "',"
                        ZSql = ZSql & "'" + WLote3 + "',"
                        ZSql = ZSql & "'" + WTeorico + "',"
                        ZSql = ZSql & "'" + ZMetodo + "',"
                        ZSql = ZSql & "'" + ZEfluentes + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesI + "',"
                        ZSql = ZSql & "'" + ZDesEfluentesII + "',"
                        ZSql = ZSql & "'" + ZVersionI + "',"
                        ZSql = ZSql & "'" + ZVersionII + "',"
                        ZSql = ZSql & "'" + ZVersionIII + "',"
                        ZSql = ZSql & "'" + Equipo.Text + "',"
                        ZSql = ZSql & "'" + ZZMetodo(1) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(2) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(3) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(4) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(5) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(6) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(7) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(8) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(9) + "',"
                        ZSql = ZSql & "'" + ZZMetodo(10) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(1) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(2) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(3) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(4) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(5) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(6) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(7) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(8) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(9) + "',"
                        ZSql = ZSql & "'" + ZZEspecificacion(10) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(1) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(2) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(3) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(4) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(5) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(6) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(7) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(8) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(9) + "',"
                        ZSql = ZSql & "'" + ZZDesEnsayo(10) + "')"
                        spImpreHoja = ZSql
                        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
                        ZLinea = ZLinea + 1
        
                    End If
  
                End If
            
            Next A
            
            XLinea = ZLinea

            For Ciclo = XLinea To 13
            
                Linea = Linea + 1
                WLinea = Str$(Linea)
                        
                WArticulo1 = ""
                WArticulo2 = ""
                WCantidad = ""
                WCanti1 = ""
                WLote1 = ""
                WCanti2 = ""
                WLote2 = ""
                WCanti3 = ""
                WLote3 = ""
            
                Sql1 = "INSERT INTO ImpreHoja ("
                Sql2 = "Hoja ,"
                Sql3 = "Renglon ,"
                Sql4 = "Fecha ,"
                Sql5 = "Codigo1 ,"
                Sql6 = "Codigo2 ,"
                Sql7 = "Maquina ,"
                Sql8 = "Articulo1 ,"
                Sql9 = "Articulo2 ,"
                Sql10 = "Cantidad ,"
                Sql11 = "Canti1 ,"
                Sql12 = "Lote1 ,"
                Sql13 = "Canti2 ,"
                Sql14 = "Lote2 ,"
                Sql15 = "Canti3 ,"
                Sql16 = "Lote3 ,"
                Sql17 = "Teorico )"
                Sql18 = "Values ("
                Sql19 = "'" + WHoja + "',"
                Sql20 = "'" + WLinea + "',"
                Sql21 = "'" + WFecha + "',"
                Sql22 = "'" + WCodigo1 + "',"
                Sql23 = "'" + WCodigo2 + "',"
                Sql24 = "'" + WMaquina + "',"
                Sql25 = "'" + WArticulo1 + "',"
                Sql26 = "'" + WArticulo2 + "',"
                Sql27 = "'" + WCantidad + "',"
                Sql28 = "'" + WCanti1 + "',"
                Sql29 = "'" + WLote1 + "',"
                Sql30 = "'" + WCanti2 + "',"
                Sql31 = "'" + WLote2 + "',"
                Sql32 = "'" + WCanti3 + "',"
                Sql33 = "'" + WLote3 + "',"
                Sql34 = "'" + WTeorico + "')"
        
                spImpreHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                        Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                        Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                        Sql31 + Sql32 + Sql33 + Sql34
                Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo

            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.ReportFileName = "ImpreHojaPelli.rpt"
            Listado.GroupSelectionFormula = "{ImpreHoja.Hoja} in 0 to 999999"
            Listado.SQLQuery = "SELECT ImpreHoja.Hoja, ImpreHoja.Fecha, ImpreHoja.Codigo1, ImpreHoja.Codigo2, ImpreHoja.Maquina, ImpreHoja.Articulo1, ImpreHoja.Articulo2, ImpreHoja.Cantidad, ImpreHoja.Canti1, ImpreHoja.Lote1, ImpreHoja.Canti2, ImpreHoja.Lote2, ImpreHoja.CAnti3, ImpreHoja.Lote3, ImpreHoja.Teorico, ImpreHoja.Renglon " _
                + "From " _
                + DSQ + ".dbo.ImpreHoja ImpreHoja " _
                + "Where " _
                + "ImpreHoja.Hoja >= 0 AND " _
                + "ImpreHoja.Hoja <= 999999"
                
            Listado.Destination = 1
            Rem Listado.Destination = 0
            Listado.Connect = Connect()
            Listado.Action = 1
           
    End Select

End Sub


Sub LeeDatos()

    ZZMetodo(1) = ""
    ZZEspecificacion(1) = ""
    ZZMetodo(2) = ""
    ZZEspecificacion(2) = ""
    ZZMetodo(3) = ""
    ZZEspecificacion(3) = ""
    ZZMetodo(4) = ""
    ZZEspecificacion(4) = ""
    ZZMetodo(5) = ""
    ZZEspecificacion(5) = ""
    ZZMetodo(6) = ""
    ZZEspecificacion(6) = ""
    ZZMetodo(7) = ""
    ZZEspecificacion(7) = ""
    ZZMetodo(8) = ""
    ZZEspecificacion(8) = ""
    ZZMetodo(9) = ""
    ZZEspecificacion(9) = ""
    ZZMetodo(10) = ""
    ZZEspecificacion(10) = ""
    
    ZZEspecificacion(11) = ""
    ZZEspecificacion(12) = ""
    ZZEspecificacion(13) = ""
    ZZEspecificacion(14) = ""
    ZZEspecificacion(15) = ""
    ZZEspecificacion(16) = ""
    ZZEspecificacion(17) = ""
    ZZEspecificacion(18) = ""
    ZZEspecificacion(19) = ""
    ZZEspecificacion(20) = ""
    
    ZZDesEnsayo(1) = ""
    ZZDesEnsayo(2) = ""
    ZZDesEnsayo(3) = ""
    ZZDesEnsayo(4) = ""
    ZZDesEnsayo(5) = ""
    ZZDesEnsayo(6) = ""
    ZZDesEnsayo(7) = ""
    ZZDesEnsayo(8) = ""
    ZZDesEnsayo(9) = ""
    ZZDesEnsayo(10) = ""
    
    Erase ZZEnsayo, ZZStd
    
    Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,ensayo4,ensayo5,Ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
            
        ZZEnsayo(1) = rstEspecifUnifica!Ensayo1
        ZZEnsayo(2) = rstEspecifUnifica!Ensayo2
        ZZEnsayo(3) = rstEspecifUnifica!Ensayo3
        ZZEnsayo(4) = rstEspecifUnifica!Ensayo4
        ZZEnsayo(5) = rstEspecifUnifica!Ensayo5
        ZZEnsayo(6) = rstEspecifUnifica!Ensayo6
        ZZEnsayo(7) = rstEspecifUnifica!Ensayo7
        ZZEnsayo(8) = rstEspecifUnifica!Ensayo8
        ZZEnsayo(9) = rstEspecifUnifica!Ensayo9
        ZZEnsayo(10) = rstEspecifUnifica!Ensayo10
                            
        ZZStd(1, 1) = rstEspecifUnifica!Valor1
        ZZStd(2, 1) = rstEspecifUnifica!valor2
        ZZStd(3, 1) = rstEspecifUnifica!Valor3
        ZZStd(4, 1) = rstEspecifUnifica!valor4
        ZZStd(5, 1) = rstEspecifUnifica!valor5
        ZZStd(6, 1) = rstEspecifUnifica!valor6
        ZZStd(7, 1) = rstEspecifUnifica!valor7
        ZZStd(8, 1) = rstEspecifUnifica!valor8
        ZZStd(9, 1) = rstEspecifUnifica!valor9
        ZZStd(10, 1) = rstEspecifUnifica!valor10
                            
        ZZStd(1, 2) = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
        ZZStd(2, 2) = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
        ZZStd(3, 2) = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
        ZZStd(4, 2) = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
        ZZStd(5, 2) = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
        ZZStd(6, 2) = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
        ZZStd(7, 2) = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
        ZZStd(8, 2) = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
        ZZStd(9, 2) = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
        ZZStd(10, 2) = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
        rstEspecifUnifica.Close
    End If
        
    Sql1 = "Select desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Valor1Ing,Valor2Ing,Valor3Ing,Valor4Ing,Valor5Ing,Valor6Ing,Valor7Ing,Valor8Ing,Valor9Ing,Valor10Ing,Valor11Ing,Valor22Ing,Valor33Ing,Valor44Ing,Valor55Ing,Valor66Ing,Valor77Ing,Valor88Ing,Valor99Ing,Valor1010Ing,Version"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
        
        
        ZZStd(1, 3) = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
        ZZStd(2, 3) = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
        ZZStd(3, 3) = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
        ZZStd(4, 3) = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
        ZZStd(5, 3) = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
        ZZStd(6, 3) = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
        ZZStd(7, 3) = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
        ZZStd(8, 3) = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
        ZZStd(9, 3) = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
        ZZStd(10, 3) = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                
        ZZStd(1, 4) = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
        ZZStd(2, 4) = IIf(IsNull(rstEspecifUnifica!HAsta2), "", rstEspecifUnifica!HAsta2)
        ZZStd(3, 4) = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
        ZZStd(4, 4) = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
        ZZStd(5, 4) = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
        ZZStd(6, 4) = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
        ZZStd(7, 4) = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
        ZZStd(8, 4) = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
        ZZStd(9, 4) = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
        ZZStd(10, 4) = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                            
        ZZStd(1, 5) = IIf(IsNull(rstEspecifUnifica!Valor1Ing), "", rstEspecifUnifica!Valor1Ing)
        ZZStd(2, 5) = IIf(IsNull(rstEspecifUnifica!Valor2Ing), "", rstEspecifUnifica!Valor2Ing)
        ZZStd(3, 5) = IIf(IsNull(rstEspecifUnifica!Valor3Ing), "", rstEspecifUnifica!Valor3Ing)
        ZZStd(4, 5) = IIf(IsNull(rstEspecifUnifica!Valor4Ing), "", rstEspecifUnifica!Valor4Ing)
        ZZStd(5, 5) = IIf(IsNull(rstEspecifUnifica!Valor5Ing), "", rstEspecifUnifica!Valor5Ing)
        ZZStd(6, 5) = IIf(IsNull(rstEspecifUnifica!Valor6Ing), "", rstEspecifUnifica!Valor6Ing)
        ZZStd(7, 5) = IIf(IsNull(rstEspecifUnifica!Valor7Ing), "", rstEspecifUnifica!Valor7Ing)
        ZZStd(8, 5) = IIf(IsNull(rstEspecifUnifica!Valor8Ing), "", rstEspecifUnifica!Valor8Ing)
        ZZStd(9, 5) = IIf(IsNull(rstEspecifUnifica!Valor9Ing), "", rstEspecifUnifica!Valor9Ing)
        ZZStd(10, 5) = IIf(IsNull(rstEspecifUnifica!Valor10Ing), "", rstEspecifUnifica!Valor10Ing)
                            
        ZZStd(1, 6) = IIf(IsNull(rstEspecifUnifica!Valor11Ing), "", rstEspecifUnifica!Valor11Ing)
        ZZStd(2, 6) = IIf(IsNull(rstEspecifUnifica!Valor22Ing), "", rstEspecifUnifica!Valor22Ing)
        ZZStd(3, 6) = IIf(IsNull(rstEspecifUnifica!Valor33Ing), "", rstEspecifUnifica!Valor33Ing)
        ZZStd(4, 6) = IIf(IsNull(rstEspecifUnifica!Valor44Ing), "", rstEspecifUnifica!Valor44Ing)
        ZZStd(5, 6) = IIf(IsNull(rstEspecifUnifica!Valor55Ing), "", rstEspecifUnifica!Valor55Ing)
        ZZStd(6, 6) = IIf(IsNull(rstEspecifUnifica!Valor66Ing), "", rstEspecifUnifica!Valor66Ing)
        ZZStd(7, 6) = IIf(IsNull(rstEspecifUnifica!Valor77Ing), "", rstEspecifUnifica!Valor77Ing)
        ZZStd(8, 6) = IIf(IsNull(rstEspecifUnifica!Valor88Ing), "", rstEspecifUnifica!Valor88Ing)
        ZZStd(9, 6) = IIf(IsNull(rstEspecifUnifica!Valor99Ing), "", rstEspecifUnifica!Valor99Ing)
        ZZStd(10, 6) = IIf(IsNull(rstEspecifUnifica!Valor1010Ing), "", rstEspecifUnifica!Valor1010Ing)
                            
        rstEspecifUnifica.Close
        
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(1) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(2) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(3) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(4) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(5) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(6) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(7) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(8) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(9) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZEnsayo(10) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDescriII(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
    End If
    
    
    For CiclaMetodo = 1 To 10
    
        If Val(ZZEnsayo(CiclaMetodo)) <> 0 Then
        
            If Val(ZZStd(CiclaMetodo, 3)) <> 0 Or Val(ZZStd(CiclaMetodo, 4)) <> 0 Then
                ZValorNormalI = Trim(ZZStd(CiclaMetodo, 3)) + " - " + Trim(ZZStd(CiclaMetodo, 4)) + " " + Trim(ZZDescriII(CiclaMetodo)) + " " + Left$(ZZStd(CiclaMetodo, 1), 50)
                ZValorNormalII = Left$(ZZStd(CiclaMetodo, 2), 50)
                    Else
                ZValorNormalI = Left$(ZZStd(CiclaMetodo, 1), 50)
                ZValorNormalII = Left$(ZZStd(CiclaMetodo, 2), 50)
            End If
        
            ZZMetodo(CiclaMetodo) = ZZEnsayo(CiclaMetodo)
            ZZEspecificacion(CiclaMetodo) = ZValorNormalI
            ZZEspecificacion(CiclaMetodo + 10) = ZValorNormalII
            
        End If
                    
    Next CiclaMetodo
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(1) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(1) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(2) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(2) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(3) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(3) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(4) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(4) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(5) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(5) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(6) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(6) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(7) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(7) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(8) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(8) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(9) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(9) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    spEnsayo = "ConsultaEnsayos " + "'" + ZZMetodo(10) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        ZZDesEnsayo(10) = rstEnsayo!Descripcion
        rstEnsayo.Close
    End If

    ZZDesEnsayo(1) = StrConv(ZZDesEnsayo(1), 2)
    ZZDesEnsayo(2) = StrConv(ZZDesEnsayo(2), 2)
    ZZDesEnsayo(3) = StrConv(ZZDesEnsayo(3), 2)
    ZZDesEnsayo(4) = StrConv(ZZDesEnsayo(4), 2)
    ZZDesEnsayo(5) = StrConv(ZZDesEnsayo(5), 2)
    ZZDesEnsayo(6) = StrConv(ZZDesEnsayo(6), 2)
    ZZDesEnsayo(7) = StrConv(ZZDesEnsayo(7), 2)
    ZZDesEnsayo(8) = StrConv(ZZDesEnsayo(8), 2)
    ZZDesEnsayo(9) = StrConv(ZZDesEnsayo(9), 2)
    ZZDesEnsayo(10) = StrConv(ZZDesEnsayo(10), 2)
    
    ZZEspecificacion(1) = StrConv(ZZEspecificacion(1), 2)
    ZZEspecificacion(2) = StrConv(ZZEspecificacion(2), 2)
    ZZEspecificacion(3) = StrConv(ZZEspecificacion(3), 2)
    ZZEspecificacion(4) = StrConv(ZZEspecificacion(4), 2)
    ZZEspecificacion(5) = StrConv(ZZEspecificacion(5), 2)
    ZZEspecificacion(6) = StrConv(ZZEspecificacion(6), 2)
    ZZEspecificacion(7) = StrConv(ZZEspecificacion(7), 2)
    ZZEspecificacion(8) = StrConv(ZZEspecificacion(8), 2)
    ZZEspecificacion(9) = StrConv(ZZEspecificacion(9), 2)
    ZZEspecificacion(10) = StrConv(ZZEspecificacion(10), 2)
    ZZEspecificacion(11) = StrConv(ZZEspecificacion(11), 2)
    ZZEspecificacion(12) = StrConv(ZZEspecificacion(12), 2)
    ZZEspecificacion(13) = StrConv(ZZEspecificacion(13), 2)
    ZZEspecificacion(14) = StrConv(ZZEspecificacion(14), 2)
    ZZEspecificacion(15) = StrConv(ZZEspecificacion(15), 2)
    ZZEspecificacion(16) = StrConv(ZZEspecificacion(16), 2)
    ZZEspecificacion(17) = StrConv(ZZEspecificacion(17), 2)
    ZZEspecificacion(18) = StrConv(ZZEspecificacion(18), 2)
    ZZEspecificacion(19) = StrConv(ZZEspecificacion(19), 2)
    ZZEspecificacion(20) = StrConv(ZZEspecificacion(20), 2)

End Sub






Private Sub Impresion_Farma()

    ZSql = "DELETE ImpreCarga"
    spImpreCarga = ZSql
    Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = "DELETE ImpreCargaI"
    spImpreCargaI = ZSql
    Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase ZImpreCarga
    ZRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "1"
                    ZImpreCarga(ZRenglon, 2) = rstCargaI!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = rstCargaI!WDescripcionII
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, MaterialAuxiliar.Descripcion as [WDescripcion]"
    ZSql = ZSql + " FROM CargaII, MaterialAuxiliar"
    ZSql = ZSql + " Where CargaII.MaterialAuxiliar = MaterialAuxiliar.Codigo"
    ZSql = ZSql + " and CargaII.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaII.Clave"
    
    rsCargaII = ZSql
    Set rstCargaII = db.OpenRecordset(rsCargaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaII.RecordCount > 0 Then
        With rstCargaII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "2"
                    ZImpreCarga(ZRenglon, 2) = rstCargaII!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = ""
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaII.Close
    End If
    
    
    
    ZLugarHumedad = 0
    Erase ZHumedad
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " and CargaIII.Humedad = 1"
    ZSql = ZSql + " and CargaIII.Renglon = 1"
    ZSql = ZSql + " Order by CargaIII.Clave"
    
    rscargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rscargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        With rstCargaIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugarHumedad = ZLugarHumedad + 1
                    ZHumedad(ZLugarHumedad) = !Equipo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIII.Close
    End If
    
    
    
    
    
    
    
    
    
    For ZCiclo = 1 To 100
        
        ZTipo = ZImpreCarga(ZCiclo, 1)
        ZDescripcion = ZImpreCarga(ZCiclo, 2)
        ZDescripcionII = ZImpreCarga(ZCiclo, 3)
        
        If ZTipo <> "" Then
                                
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ImpreCarga ("
            ZSql = ZSql & "Partida ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Terminado ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Tipo ,"
            ZSql = ZSql & "DescripcionI ,"
            ZSql = ZSql & "DescripcionII )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + Hoja.Text + "',"
            ZSql = ZSql & "'" + DesProducto.Caption + "',"
            ZSql = ZSql & "'" + Producto.Text + "',"
            ZSql = ZSql & "'" + Teorico.Text + "',"
            ZSql = ZSql & "'" + ZTipo + "',"
            ZSql = ZSql & "'" + ZDescripcion + "',"
            ZSql = ZSql & "'" + ZDescripcionII + "')"
        
            spImpreCarga = ZSql
            Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
                            
    Next ZCiclo
    
    
    
    
    
    
    
    
    
    
    Erase ZImpreMetodo
    ZLugarMetodo = 0
    Erase ZImpreCargaI
    ZRenglon = 0
    ZLugar = 1
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZEquipo = rstCargaI!Equipo
                    ZZDescripcionI = rstCargaI!WDescripcion
                    ZZDescripcionII = rstCargaI!WDescripcionII
                    ZZZMetodo = IIf(IsNull(rstCargaI!Metodo), "", rstCargaI!Metodo)
                    ZZCantidad = IIf(IsNull(rstCargaI!Cantidad), "0", rstCargaI!Cantidad)
                    
                    If Trim(ZZZMetodo) <> "" Then
                    
                        ZEntraMetodo = "S"
                    
                        For ZCicloMetodo = 1 To ZLugarMetodo
                            If ZImpreMetodo(ZCicloMetodo) = Trim(ZZZMetodo) Then
                                ZEntraMetodo = "N"
                                Exit For
                            End If
                        Next ZCicloMetodo
                    
                        If ZEntraMetodo = "S" Then
                            ZLugarMetodo = ZLugarMetodo + 1
                            ZImpreMetodo(ZLugarMetodo) = Trim(ZZZMetodo)
                        End If
                        
                    End If
                    
                    For Ciclo = 1 To ZZCantidad
                        Select Case ZLugar
                            Case 1
                                ZRenglon = ZRenglon + 1
                                ZImpreCargaI(ZRenglon, 1) = Producto.Text
                                ZImpreCargaI(ZRenglon, 2) = DesProducto.Caption
                                ZImpreCargaI(ZRenglon, 3) = Hoja.Text
                                ZImpreCargaI(ZRenglon, 4) = Teorico.Text
                                
                                ZImpreCargaI(ZRenglon, 5) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 6) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 7) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 8) = ZZZMetodo
                                
                                ZLugar = 2
                            
                            Case 2
                                ZImpreCargaI(ZRenglon, 9) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 10) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 11) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 12) = ZZZMetodo
                            
                                ZLugar = 3
                                
                            Case 3
                                ZImpreCargaI(ZRenglon, 13) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 14) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 15) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 16) = ZZZMetodo
                                
                                ZLugar = 1
                            Case Else
                        End Select
                    Next Ciclo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    
    For ZCiclo = 1 To ZRenglon
    
        ZZCodigo = Str$(ZCiclo)
        
        ZZTerminado = ZImpreCargaI(ZCiclo, 1)
        ZZDescripcion = ZImpreCargaI(ZCiclo, 2)
        ZZPartida = ZImpreCargaI(ZCiclo, 3)
        ZZCantidad = ZImpreCargaI(ZCiclo, 4)
                                
        ZZEquipoI = ZImpreCargaI(ZCiclo, 5)
        ZZDesEquipoI = ZImpreCargaI(ZCiclo, 6)
        ZZDesEquipoOtroI = ZImpreCargaI(ZCiclo, 7)
        ZZMetodoI = ZImpreCargaI(ZCiclo, 8)
                                
        ZZEquipoII = ZImpreCargaI(ZCiclo, 9)
        ZZDesEquipoII = ZImpreCargaI(ZCiclo, 10)
        ZZDesEquipoOtroII = ZImpreCargaI(ZCiclo, 11)
        ZZMetodoII = ZImpreCargaI(ZCiclo, 12)
                                
        ZZEquipoIII = ZImpreCargaI(ZCiclo, 13)
        ZZDesEquipoIII = ZImpreCargaI(ZCiclo, 14)
        ZZDesEquipoOtroIII = ZImpreCargaI(ZCiclo, 15)
        ZZMetodoIII = ZImpreCargaI(ZCiclo, 16)
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreCargaI ("
        ZSql = ZSql & "Codigo ,"
        ZSql = ZSql & "Terminado ,"
        ZSql = ZSql & "Descripcion ,"
        ZSql = ZSql & "Partida ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "MetodoI ,"
        ZSql = ZSql & "EquipoI ,"
        ZSql = ZSql & "DesEquipoI ,"
        ZSql = ZSql & "DesEquipoOtroI ,"
        ZSql = ZSql & "MetodoII ,"
        ZSql = ZSql & "EquipoII ,"
        ZSql = ZSql & "DesEquipoII ,"
        ZSql = ZSql & "DesEquipoOtroII ,"
        ZSql = ZSql & "MetodoIII ,"
        ZSql = ZSql & "EquipoIII ,"
        ZSql = ZSql & "DesEquipoIII ,"
        ZSql = ZSql & "DesEquipoOtroIII )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + ZZCodigo + "',"
        ZSql = ZSql & "'" + ZZTerminado + "',"
        ZSql = ZSql & "'" + ZZDescripcion + "',"
        ZSql = ZSql & "'" + ZZPartida + "',"
        ZSql = ZSql & "'" + Str$(ZZCantidad) + "',"
        ZSql = ZSql & "'" + ZZMetodoI + "',"
        ZSql = ZSql & "'" + ZZEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroI + "',"
        ZSql = ZSql & "'" + ZZMetodoII + "',"
        ZSql = ZSql & "'" + ZZEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroII + "',"
        ZSql = ZSql & "'" + ZZMetodoIII + "',"
        ZSql = ZSql & "'" + ZZEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroIII + "')"
        
        spImpreCargaI = ZSql
        Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
                            
    Next ZCiclo
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " Partida = " + "'" + Hoja.Text + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + Teorico.Text + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Producto.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " Partida = " + "'" + Hoja.Text + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + Teorico.Text + "',"
    ZSql = ZSql + " ImprePaso = Paso "
    ZSql = ZSql + " Where Terminado = " + "'" + Producto.Text + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    

    Producto.Text = UCase(Producto.Text)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    
    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, " _
                + "Terminado.Descripcion, Terminado.Version, Terminado.FechaVersion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

    Listado.ReportFileName = "ImpreCaratula.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    
    Listado.SQLQuery = "SELECT ImpreCarga.Partida, ImpreCarga.Terminado, ImpreCarga.Descripcion, ImpreCarga.Cantidad, ImpreCarga.Tipo, ImpreCarga.DescripcionI, ImpreCarga.DescripcionII " _
                + "From " _
                + DSQ + ".dbo.ImpreCarga ImpreCarga " _
                + "Where " _
                + "ImpreCarga.Partida >= " + Hoja.Text + " AND " _
                + "ImpreCarga.Partida <= " + Hoja.Text

    Listado.ReportFileName = "ImpreEquipos.rpt"
    Listado.GroupSelectionFormula = "{ImpreCarga.Partida} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{ImpreCarga.Partida} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    

    
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Articulo, CargaIII.PTerminado, CargaIII.Letra, CargaIII.Descripcion, CargaIII.Cantidad, CargaIII.Partida, CargaIII.CantidadPartida , " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaIII CargaIII, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaIII.Terminado = Terminado.Codigo AND " _
                    + "CargaIII.Terminado >= '" + Producto.Text + "' AND " _
                    + "CargaIII.Terminado <= '" + Producto.Text + "'"

    Listado.ReportFileName = "ImpreProcedimiento.rpt"
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.Destination = 1
    Listado.Action = 1
    
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Partida, CargaIII.CantidadPartida, CargaIII.Peso, CargaIII.ImprePeso, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaIII CargaIII, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "CargaIII.Terminado = Terminado.Codigo AND " _
                + "CargaIII.Terminado >= '" + Producto.Text + "' AND " _
                + "CargaIII.Terminado <= '" + Producto.Text + "' AND " _
                + "CargaIII.Peso = 1 AND " _
                + "CargaIII.ImprePeso = 'S'"

    Listado.ReportFileName = "ImprePeso.rpt"
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.Destination = 1
    Listado.Action = 1


    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
            + "Terminado.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Hoja Hoja, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "Hoja.Producto = Terminado.Codigo AND " _
            + "Hoja.Hoja >= " + Hoja.Text + " AND " _
            + "Hoja.Hoja <= " + Hoja.Text
            
    Listado.ReportFileName = "ImpreObservaciones.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    Listado.Action = 1
    
    
    
    
    
    






    For A = 1 To 40
    
        Tipo = UCase(Grilla.TextMatrix(A, 1))
        Terminado = UCase(Grilla.TextMatrix(A, 2))
        Articulo = UCase(Grilla.TextMatrix(A, 3))
        ImpreArticulo = UCase(Grilla.TextMatrix(A, 4))
        Cantidad = Grilla.TextMatrix(A, 5)
                
        If Articulo <> "" Then
                
            Auxi = Str$(A)
            Call Ceros(Auxi, 2)
                
            Auxi1 = Str$(Hoja.Text)
            Call Ceros(Auxi1, 6)
            
            WClave = Auxi1 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " ImpreArticulo = " + "'" + ImpreArticulo + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next A





    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.ImpreArticulo," _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

    Listado.ReportFileName = "ImpreHojaFarmaAlmacen.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    
    
    
    For CicloHumedad = 1 To ZLugarHumedad
    
        WIdentificacion = ""
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + ZHumedad(CicloHumedad) + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            WIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            rstEquipo.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Identificacion = " + "'" + WIdentificacion + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

        Listado.ReportFileName = "ImpreHumedad.rpt"
        Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
        Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
        Listado.Destination = 1
        Listado.Action = 1
        
    Next CicloHumedad
    
    
    
    
    Rem Listado.SQLQuery = "SELECT ImpreCargaI.Codigo, ImpreCargaI.Terminado, ImpreCargaI.Descripcion, ImpreCargaI.Partida, ImpreCargaI.Cantidad, ImpreCargaI.MetodoI, ImpreCargaI.EquipoI, ImpreCargaI.DesEquipoI, ImpreCargaI.DesEquipoOtroI, ImpreCargaI.MetodoII, ImpreCargaI.EquipoII, ImpreCargaI.DesEquipoII, ImpreCargaI.DesEquipoOtroII, ImpreCargaI.MetodoIII, ImpreCargaI.EquipoIII, ImpreCargaI.DesEquipoIII, ImpreCargaI.DesEquipoOtroIII " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ImpreCargaI ImpreCargaI " _
    rem             + "Where " _
    rem             + "ImpreCargaI.Codigo >= 0 AND " _
    rem             + "ImpreCargaI.Codigo <= 999999"
    Rem
    Rem Listado.ReportFileName = "ImpreIdentificacion.rpt"
    Rem Listado.GroupSelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Listado.SelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Listado.Destination = 1
    Rem Listado.Action = 1
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Listado.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Valor, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaV CargaV, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaV.Terminado = Terminado.Codigo AND " _
                    + "CargaV.Terminado >= '" + Producto.Text + "' AND " _
                    + "CargaV.Terminado <= '" + Producto.Text + "' AND " _
                    + "CargaV.Paso = 99"

    Listado.ReportFileName = "ImpreCalidad.rpt"
    
    Uno = "{CargaV.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Dos = " and {CargaV.Paso} = 99"
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    Rem For ZCiclo = 1 To ZLugarMetodo
    Rem
    Rem     XMetodo = ZImpreMetodo(ZCiclo)
    Rem
    Rem     ZSql = ""
    Rem     ZSql = ZSql + "UPDATE Lavado SET "
    Rem     ZSql = ZSql + " Terminado = " + "'" + Producto.Text + "',"
    Rem     ZSql = ZSql + " DesTerminado = " + "'" + DesProducto.Caption + "',"
    Rem     ZSql = ZSql + " Partida = " + "'" + Hoja.Text + "',"
    Rem     ZSql = ZSql + " Cantidad = " + "'" + Teorico.Text + "'"
    Rem     ZSql = ZSql + " Where Codigo = " + "'" + XMetodo + "'"
    Rem     spLavado = ZSql
    Rem     Set rstLavado = db.OpenRecordset(spLavado, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem
    Rem     Listado.SQLQuery = "SELECT Lavado.Clave, Lavado.Codigo, Lavado.Descripcion, Lavado.Terminado, Lavado.Partida, Lavado.Cantidad " _
    rem                     + "From " _
    rem                     + DSQ + ".dbo.Lavado Lavado " _
    rem                     + "Where " _
    rem                     + "Lavado.Codigo >= '" + XMetodo + "' AND " _
    rem                     + "Lavado.Codigo <= '" + XMetodo + "'"
    Rem
    Rem     Listado.ReportFileName = "ImpreLavado.rpt"
    Rem     Listado.GroupSelectionFormula = "{Lavado.Codigo} in " + Chr$(34) + XMetodo + Chr$(34) + " to " + Chr$(34) + XMetodo + Chr$(34)
    Rem     Listado.SelectionFormula = "{Lavado.Codigo} in " + Chr$(34) + XMetodo + Chr$(34) + " to " + Chr$(34) + XMetodo + Chr$(34)
    Rem     Listado.Destination = 1
    Rem     Listado.Action = 1
    Rem
    Rem Next ZCiclo

End Sub


Private Sub Impresion_FarmaII()

    Producto.Text = UCase(Producto.Text)

    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZDesTerminado = rstTerminado!Descripcion
        rstTerminado.Close
    End If

    ZSql = "DELETE ImpreCarga"
    spImpreCarga = ZSql
    Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = "DELETE ImpreCargaI"
    spImpreCargaI = ZSql
    Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Erase ZImpreCarga
    ZRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII], Equipo.Poe as [WPoe], Equipo.Identificacion as [WIdentificacion], Equipo.PoeLimpieza as [WPoeLimpieza]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "1"
                    ZImpreCarga(ZRenglon, 2) = rstCargaI!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = rstCargaI!WDescripcionII
                    ZImpreCarga(ZRenglon, 4) = rstCargaI!WPoe
                    ZImpreCarga(ZRenglon, 5) = rstCargaI!WIdentificacion
                    ZImpreCarga(ZRenglon, 6) = rstCargaI!WPoeLimpieza
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, MaterialAuxiliar.Descripcion as [WDescripcion]"
    ZSql = ZSql + " FROM CargaII, MaterialAuxiliar"
    ZSql = ZSql + " Where CargaII.MaterialAuxiliar = MaterialAuxiliar.Codigo"
    ZSql = ZSql + " and CargaII.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaII.Clave"
    
    rsCargaII = ZSql
    Set rstCargaII = db.OpenRecordset(rsCargaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaII.RecordCount > 0 Then
        With rstCargaII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "2"
                    ZImpreCarga(ZRenglon, 2) = rstCargaII!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = ""
                    ZImpreCarga(ZRenglon, 4) = ""
                    ZImpreCarga(ZRenglon, 5) = ""
                    ZImpreCarga(ZRenglon, 6) = ""
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaII.Close
    End If
    
    
    
    ZLugarHumedad = 0
    Erase ZHumedad
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " and CargaIII.Humedad = 1"
    ZSql = ZSql + " and CargaIII.Renglon = 1"
    ZSql = ZSql + " Order by CargaIII.Clave"
    
    rscargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rscargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        With rstCargaIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugarHumedad = ZLugarHumedad + 1
                    ZHumedad(ZLugarHumedad) = !Equipo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIII.Close
    End If
    
    
    
    
    
    
    
    
    
    For ZCiclo = 1 To 100
        
        ZTipo = ZImpreCarga(ZCiclo, 1)
        ZDescripcion = ZImpreCarga(ZCiclo, 2)
        ZDescripcionII = ZImpreCarga(ZCiclo, 3)
        ZPoe = Trim(ZImpreCarga(ZCiclo, 4))
        ZIdentificacion = Trim(ZImpreCarga(ZCiclo, 5))
        ZPoeLimpieza = Trim(ZImpreCarga(ZCiclo, 6))
        If ZIdentificacion <> "" Then
            ZDescripcion = ZIdentificacion + " - " + ZDescripcion
        End If
        
        
        If ZTipo <> "" Then
                                
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ImpreCarga ("
            ZSql = ZSql & "Partida ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Terminado ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Tipo ,"
            ZSql = ZSql & "DescripcionI ,"
            ZSql = ZSql & "DescripcionII )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + Hoja.Text + "',"
            ZSql = ZSql & "'" + ZDesTerminado + "',"
            ZSql = ZSql & "'" + Producto.Text + "',"
            ZSql = ZSql & "'" + Teorico.Text + "',"
            ZSql = ZSql & "'" + ZTipo + "',"
            ZSql = ZSql & "'" + ZDescripcion + "',"
            ZSql = ZSql & "'" + ZDescripcionII + "')"
        
            spImpreCarga = ZSql
            Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
                            
    Next ZCiclo
    
    
    
    
    
    
    
    
    
    
    Erase ZImpreMetodo
    ZLugarMetodo = 0
    Erase ZImpreCargaI
    ZRenglon = 0
    ZLugar = 1
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII], Equipo.Poe as [WPoe], Equipo.Identificacion as [WIdentificacion], Equipo.PoeLimpieza as [WPoeLimpieza]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZEquipo = rstCargaI!Equipo
                    ZZDescripcionI = rstCargaI!WDescripcion
                    ZZDescripcionII = rstCargaI!WDescripcionII
                    ZZZMetodo = Trim(rstCargaI!WPoeLimpieza) + " - " + IIf(IsNull(rstCargaI!Metodo), "", rstCargaI!Metodo)
                    ZZCantidad = IIf(IsNull(rstCargaI!Cantidad), "0", rstCargaI!Cantidad)
                    ZZPoe = rstCargaI!WPoe
                    ZZIdentificacion = rstCargaI!WIdentificacion
                    ZZPoeLimpieza = rstCargaI!WPoeLimpieza
                    If Trim(ZZIdentificacion) <> "" Then
                        ZZDescripcionI = Trim(ZZIdentificacion) + " - " + ZZDescripcionI
                    End If
                    
                    Rem If ZZCantidad <> 0 Then
                    Rem
                    Rem     ZEntraMetodo = "S"
                    Rem
                    Rem     For ZCicloMetodo = 1 To ZLugarMetodo
                    Rem         If ZImpreMetodo(ZCicloMetodo) = Trim(ZZZMetodo) Then
                    Rem             ZEntraMetodo = "N"
                    Rem             Exit For
                    Rem         End If
                    Rem     Next ZCicloMetodo
                    Rem
                    Rem     If ZEntraMetodo = "S" Then
                    Rem         ZLugarMetodo = ZLugarMetodo + 1
                    Rem         ZImpreMetodo(ZLugarMetodo) = Trim(ZZZMetodo)
                    Rem     End If
                    Rem
                    Rem End If
                    
                    For Ciclo = 1 To ZZCantidad
                        Select Case ZLugar
                            Case 1
                                ZRenglon = ZRenglon + 1
                                ZImpreCargaI(ZRenglon, 1) = Producto.Text
                                ZImpreCargaI(ZRenglon, 2) = ZDesTerminado
                                ZImpreCargaI(ZRenglon, 3) = "0"
                                ZImpreCargaI(ZRenglon, 4) = "0"
                                
                                ZImpreCargaI(ZRenglon, 5) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 6) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 7) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 8) = ZZZMetodo
                                
                                ZLugar = 2
                            
                            Case 2
                                ZImpreCargaI(ZRenglon, 9) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 10) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 11) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 12) = ZZZMetodo
                            
                                ZLugar = 3
                                
                            Case 3
                                ZImpreCargaI(ZRenglon, 13) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 14) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 15) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 16) = ZZZMetodo
                                
                                ZLugar = 1
                            Case Else
                        End Select
                    Next Ciclo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    
    For ZCiclo = 1 To ZRenglon
    
        ZZCodigo = Str$(ZCiclo)
        
        ZZTerminado = ZImpreCargaI(ZCiclo, 1)
        ZZDescripcion = ZImpreCargaI(ZCiclo, 2)
        ZZPartida = ZImpreCargaI(ZCiclo, 3)
        ZZCantidad = ZImpreCargaI(ZCiclo, 4)
                                
        ZZEquipoI = ZImpreCargaI(ZCiclo, 5)
        ZZDesEquipoI = ZImpreCargaI(ZCiclo, 6)
        ZZDesEquipoOtroI = ZImpreCargaI(ZCiclo, 7)
        ZZMetodoI = ZImpreCargaI(ZCiclo, 8)
                                
        ZZEquipoII = ZImpreCargaI(ZCiclo, 9)
        ZZDesEquipoII = ZImpreCargaI(ZCiclo, 10)
        ZZDesEquipoOtroII = ZImpreCargaI(ZCiclo, 11)
        ZZMetodoII = ZImpreCargaI(ZCiclo, 12)
                                
        ZZEquipoIII = ZImpreCargaI(ZCiclo, 13)
        ZZDesEquipoIII = ZImpreCargaI(ZCiclo, 14)
        ZZDesEquipoOtroIII = ZImpreCargaI(ZCiclo, 15)
        ZZMetodoIII = ZImpreCargaI(ZCiclo, 16)
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreCargaI ("
        ZSql = ZSql & "Codigo ,"
        ZSql = ZSql & "Terminado ,"
        ZSql = ZSql & "Descripcion ,"
        ZSql = ZSql & "Partida ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "MetodoI ,"
        ZSql = ZSql & "EquipoI ,"
        ZSql = ZSql & "DesEquipoI ,"
        ZSql = ZSql & "DesEquipoOtroI ,"
        ZSql = ZSql & "MetodoII ,"
        ZSql = ZSql & "EquipoII ,"
        ZSql = ZSql & "DesEquipoII ,"
        ZSql = ZSql & "DesEquipoOtroII ,"
        ZSql = ZSql & "MetodoIII ,"
        ZSql = ZSql & "EquipoIII ,"
        ZSql = ZSql & "DesEquipoIII ,"
        ZSql = ZSql & "DesEquipoOtroIII )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + ZZCodigo + "',"
        ZSql = ZSql & "'" + ZZTerminado + "',"
        ZSql = ZSql & "'" + ZZDescripcion + "',"
        ZSql = ZSql & "'" + ZZPartida + "',"
        ZSql = ZSql & "'" + Str$(ZZCantidad) + "',"
        ZSql = ZSql & "'" + ZZMetodoI + "',"
        ZSql = ZSql & "'" + ZZEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroI + "',"
        ZSql = ZSql & "'" + ZZMetodoII + "',"
        ZSql = ZSql & "'" + ZZEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroII + "',"
        ZSql = ZSql & "'" + ZZMetodoIII + "',"
        ZSql = ZSql & "'" + ZZEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroIII + "')"
        
        spImpreCargaI = ZSql
        Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
                            
    Next ZCiclo
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " Partida = " + "'" + Hoja.Text + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + Teorico.Text + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Producto.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " Partida = " + "'" + Hoja.Text + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + Teorico.Text + "',"
    ZSql = ZSql + " ImprePaso = Paso "
    ZSql = ZSql + " Where Terminado = " + "'" + Producto.Text + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem BY NAN BUSCO VERSION Y FECHA DE VERSION 6-6-2014*****************
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Producto.Text + "'"
    ZSql = ZSql + " Order by CargaIII.Clave"
    rscargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rscargaIII, dbOpenSnapshot, dbSQLPassThrough)
      If rstCargaIII.RecordCount > 0 Then
          vers = IIf(IsNull(rstCargaIII!Version), 0, rstCargaIII!Version)
          fechaversio = IIf(IsNull(rstCargaIII!FechaVersion), "", rstCargaIII!FechaVersion)
       End If
    rstCargaIII.Close
    
    ZSql = ""
    ZSql = ZSql + "UPDATE hoja SET "
    ZSql = ZSql + " Impreversion = " + "'" + vers + "',"
    ZSql = ZSql + " Imprefechaversion = " + "'" + fechaversio + "'"
    ZSql = ZSql + " Where hoja = " + "'" + Hoja.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Rem FIN BY NAN *******************************

    
    
    

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, " _
                + "Terminado.Descripcion, Terminado.Version, Terminado.FechaVersion, Terminado.LoteAutorizado " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

    Listado.ReportFileName = "WImpreCaratula.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    Listado.SQLQuery = "SELECT ImpreCarga.Partida, ImpreCarga.Terminado, ImpreCarga.Descripcion, ImpreCarga.Cantidad, ImpreCarga.Tipo, ImpreCarga.DescripcionI, ImpreCarga.DescripcionII " _
                + "From " _
                + DSQ + ".dbo.ImpreCarga ImpreCarga " _
                + "Where " _
                + "ImpreCarga.Partida >= " + Hoja.Text + " AND " _
                + "ImpreCarga.Partida <= " + Hoja.Text
    
    Listado.ReportFileName = "ImpreEquipos.rpt"
    Listado.GroupSelectionFormula = "{ImpreCarga.Partida} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{ImpreCarga.Partida} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    

    ZDesdePaso = "0"
    ZHastaPaso = "999"
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Renglon, CargaIII.Articulo, CargaIII.PTerminado, CargaIII.Letra, CargaIII.Descripcion, CargaIII.Cantidad, CargaIII.Partida, CargaIII.CantidadPartida , " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaIII CargaIII, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaIII.Terminado = Terminado.Codigo AND " _
                    + "CargaIII.Terminado >= '" + Producto.Text + "' AND " _
                    + "CargaIII.Terminado <= '" + Producto.Text + "' AND " _
                    + "CargaIII.Paso >= " + ZDesdePaso + " AND " _
                    + "CargaIII.Paso <= " + ZHastaPaso

    Listado.ReportFileName = "WImpreProcedimiento.rpt"
    
    Uno = "{CargaIII.Paso} in " + ZDesdePaso + " to " + ZHastaPaso
    Dos = " and {CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
        
    

    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Partida, CargaIII.CantidadPartida, CargaIII.Peso, CargaIII.ImprePeso, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaIII CargaIII, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "CargaIII.Terminado = Terminado.Codigo AND " _
                + "CargaIII.Terminado >= '" + Producto.Text + "' AND " _
                + "CargaIII.Terminado <= '" + Producto.Text + "' AND " _
                + "CargaIII.Peso = 1 AND " _
                + "CargaIII.ImprePeso = 'S'"

    Listado.ReportFileName = "ImprePeso.rpt"
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Listado.Destination = 1
    Listado.Action = 1


    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
            + "Terminado.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Hoja Hoja, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "Hoja.Producto = Terminado.Codigo AND " _
            + "Hoja.Hoja >= " + Hoja.Text + " AND " _
            + "Hoja.Hoja <= " + Hoja.Text
            
    Listado.ReportFileName = "ImpreObservaciones.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    






    For A = 1 To 40
    
        Tipo = UCase(Grilla.TextMatrix(A, 1))
        Terminado = UCase(Grilla.TextMatrix(A, 2))
        Articulo = UCase(Grilla.TextMatrix(A, 3))
        ImpreArticulo = UCase(Grilla.TextMatrix(A, 4))
        Cantidad = Grilla.TextMatrix(A, 5)
                
        If Articulo <> "" Then
                
            Auxi = Str$(A)
            Call Ceros(Auxi, 2)
                
            Auxi1 = Str$(Hoja.Text)
            Call Ceros(Auxi1, 6)
            
            WClave = Auxi1 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " ImpreArticulo = " + "'" + ImpreArticulo + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next A





    
    
    
    
    Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.ImpreArticulo," _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

    Listado.ReportFileName = "ImpreHojaFarmaAlmacen.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    
    
    
    
    
    For CicloHumedad = 1 To ZLugarHumedad
    
        WIdentificacion = ""
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + ZHumedad(CicloHumedad) + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            WIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            rstEquipo.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Identificacion = " + "'" + WIdentificacion + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + Hoja.Text + " AND " _
                + "Hoja.Hoja <= " + Hoja.Text

        Listado.ReportFileName = "ImpreHumedad.rpt"
        Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
        Listado.SelectionFormula = "{Hoja.Hoja} in " + Hoja.Text + " to " + Hoja.Text
        Listado.Destination = 1
        Listado.Action = 1
        
    Next CicloHumedad
    
    
    
    
    
    Listado.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Valor, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaV CargaV, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaV.Terminado = Terminado.Codigo AND " _
                    + "CargaV.Terminado >= '" + Producto.Text + "' AND " _
                    + "CargaV.Terminado <= '" + Producto.Text + "' AND " _
                    + "CargaV.Paso = 99"

    Listado.ReportFileName = "ImpreCalidad.rpt"
    
    Uno = "{CargaV.Terminado} in " + Chr$(34) + Producto.Text + Chr$(34) + " to " + Chr$(34) + Producto.Text + Chr$(34)
    Dos = " and {CargaV.Paso} = 99"
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Destination = 1
    Listado.Action = 1
    
    
    
    
    Rem Listado.SQLQuery = "SELECT ImpreCargaI.Codigo, ImpreCargaI.Terminado, ImpreCargaI.Descripcion, ImpreCargaI.Partida, ImpreCargaI.Cantidad, ImpreCargaI.MetodoI, ImpreCargaI.EquipoI, ImpreCargaI.DesEquipoI, ImpreCargaI.DesEquipoOtroI, ImpreCargaI.MetodoII, ImpreCargaI.EquipoII, ImpreCargaI.DesEquipoII, ImpreCargaI.DesEquipoOtroII, ImpreCargaI.MetodoIII, ImpreCargaI.EquipoIII, ImpreCargaI.DesEquipoIII, ImpreCargaI.DesEquipoOtroIII " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ImpreCargaI ImpreCargaI " _
    rem             + "Where " _
    rem             + "ImpreCargaI.Codigo >= 0 AND " _
    rem             + "ImpreCargaI.Codigo <= 999999"
    Rem
    Rem Listado.ReportFileName = "ImpreIdentificacion.rpt"
    Rem Listado.GroupSelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Listado.SelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Listado.Destination = 1
    Rem Listado.Action = 1


End Sub

Private Sub Reserva_Click()
 
    On Error GoTo WError
    
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        m$ = "La fecha de la hoja de produccion es incorrecta"
        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
        Exit Sub
    End If
    
    spHoja = "ListaHoja " + "'" + Hoja.Text + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.Close
        m$ = "Partida ya existente"
        G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
        Exit Sub
    End If
    
    Select Case Val(Wempresa)
        Case 1
            Rem If Val(Hoja.Text) > 69999 Or Val(Hoja.Text) < 57600 Then
            If Val(Hoja.Text) > 199999 Or Val(Hoja.Text) < 100000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 100000 y 199999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 2
            If Val(Hoja.Text) > 55999 Or Val(Hoja.Text) < 55300 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 55300 y 55999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 3
            Rem If Val(Hoja.Text) > 99999 Or Val(Hoja.Text) < 82000 Then
            If Val(Hoja.Text) > 299999 Or Val(Hoja.Text) < 200000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 200000 y 299999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 4
            If Val(Hoja.Text) > 19999 Or Val(Hoja.Text) < 11100 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 11100 y 19999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 5
            Rem If Val(Hoja.Text) > 9999 Or Val(Hoja.Text) < 4600 Then
            If Val(Hoja.Text) > 399999 Or Val(Hoja.Text) < 300000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 300000 y 399999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 6
            Rem If Val(Hoja.Text) > 1999 Or Val(Hoja.Text) < 1740 Then
            If Val(Hoja.Text) > 499999 Or Val(Hoja.Text) < 400000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 400000 y 499999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 7
            Rem If Val(Hoja.Text) > 999 Or Val(Hoja.Text) < 7 Then
            If Val(Hoja.Text) > 599999 Or Val(Hoja.Text) < 500000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 500000 y 599999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 8
            If Val(Hoja.Text) > 29999 Or Val(Hoja.Text) < 20800 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 20800 y 29999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 9
            If Val(Hoja.Text) > 30999 Or Val(Hoja.Text) < 30000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 30000 y 30999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 10
            If Val(Hoja.Text) > 699999 Or Val(Hoja.Text) < 600000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 600000 y 699999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case 11
            If Val(Hoja.Text) > 799999 Or Val(Hoja.Text) < 700000 Then
                m$ = "Partida fuera de rango. La misma debe estar entre el numero 700000 y 799999"
                G% = MsgBox(m$, 0, "Ingreso de Hoja de Produccion")
                Exit Sub
            End If
        Case Else
    End Select


    Renglon = 1
    Auxi = Str$(Renglon)
    Call Ceros(Auxi, 2)
                        
    Auxi1 = Str$(Hoja.Text)
    Call Ceros(Auxi1, 6)
                    
    WClave = Auxi1 + Auxi
    WHoja = Hoja.Text
    WRenglon = Str$(Renglon)
    WFecha = Fecha.Text
    WProducto = Producto.Text
    WTeorico = Teorico.Text
    WReal = "0"
    WFechaing = "  /  /    "
    WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
    WTipo = "M"
    WArticulo = "  -   -   "
    WTerminado = "  -     -   "
    WCantidad = "0"
    WLote = ""
    WDate = Date$
    WImporte = ""
    WMarca = ""
    WSaldo = "0"
    
    WLote1 = ""
    WLote2 = ""
    WLote3 = ""
    
    WCanti1 = "0"
    WCanti2 = "0"
    WCanti3 = "0"
    
    WCosto1 = "0"
    WCosto2 = "0"
    WCosto3 = "0"
                    
    XParam = "'" + WClave + "','" _
                + WHoja + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WProducto + "','" _
                + WCantidad + "','" _
                + WTipo + "','" _
                + WLote + "','" _
                + WArticulo + "','" _
                + WTerminado + "','" _
                + WTeorico + "','" _
                + WReal + "','" _
                + WFechaing + "','" _
                + WFechaingord + "','" _
                + WDate + "','" _
                + WImporte + "','" _
                + WMarca + "','" _
                + WSaldo + "','" _
                + WLote1 + "','" + WCanti1 + "','" _
                + WLote2 + "','" + WCanti2 + "','" _
                + WLote3 + "','" + WLote3 + "','" _
                + WCosto1 + "','" _
                + WCosto2 + "','" _
                + WCosto3 + "'"
                                            
    spHoja = "AltaHoja " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    VersionI.Text = "99"
    VersionII.Text = "99"
    VersionIII.Text = "99"
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " Equipo = " + "'" + Equipo.Text + "',"
    ZSql = ZSql + " VersionI = " + "'" + VersionI.Text + "',"
    ZSql = ZSql + " VersionII = " + "'" + VersionII.Text + "',"
    ZSql = ZSql + " VersionIII = " + "'" + VersionIII.Text + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + Hoja.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    
    XCodigo = Val(Mid$(Producto.Text, 4, 5))
    XTipoPro = ""
    If Val(Wempresa) = 1 Then
        If XCodigo >= 0 And XCodigo <= 999 Then
            XTipoPro = "CO"
                Else
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                XTipoPro = "CO"
                    Else
                XTipoPro = ""
            End If
        End If
    End If
    
    
    
    
    Erase ZCompo
    Renglon = 0

    spComposicion = "ConsultaComposicionProducto " + "'" + Producto.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then

                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
        
                        Renglon = Renglon + 1
                        
                        ZCompo(Renglon, 1) = rstComposicion!Tipo
                        ZCompo(Renglon, 2) = rstComposicion!Articulo1
                        ZCompo(Renglon, 3) = rstComposicion!Articulo2
                        ZCompo(Renglon, 4) = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                        ZCompo(Renglon, 5) = Str$(rstComposicion!Cantidad * Val(Teorico.Text))
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    
    For Da = 1 To Renglon
    
        Tipo = ZCompo(Da, 1)
        Auxi2 = ZCompo(Da, 2)
        Auxi1 = ZCompo(Da, 3)
        XCantidad = Val(ZCompo(Da, 4))
        
        WStock = 0
                
        Select Case Tipo
            Case "T"
                WImpre1 = Auxi1
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                    rstTerminado.Close
                End If
            Case "M"
                WImpre1 = Auxi2
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZStock = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                    rstArticulo.Close
                End If
            Case Else
        End Select
        
        If XCantidad > WStock Then
            ZCompo(Da, 4) = Str$(ZStock)
        End If
        
    Next Da
    
    
    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 2, 3, 4
            Sql1 = "DELETE ImpreHojaII"
            spImpreHojaII = Sql1
            Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
        
            WHoja = Hoja.Text
            WFecha = Fecha.Text
            WCodigo1 = Left$(Producto.Text, 2)
            WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
            Select Case Val(Wempresa)
                Case 1
                    WMaquina = "I"
                Case 2
                    WMaquina = "I"
                Case 3
                    WMaquina = "II"
                Case 4
                    WMaquina = "II"
                Case 5
                    WMaquina = "III"
                Case 6
                    WMaquina = "IV"
                Case 7
                    WMaquina = "V"
                Case 8
                    WMaquina = "V"
                Case 8
                    WMaquina = "VI"
                Case Else
            End Select
            WTeorico = Teorico.Text
            
            Linea = 0
            LineaII = 0
        
            For A = 1 To 100
        
                Tipo = ZCompo(A, 1)
                Articulo = ZCompo(A, 2)
                Terminado = ZCompo(A, 3)
                Cantidad = ZCompo(A, 4)
                CantidadII = ZCompo(A, 5)
                 
                If Tipo = "M" Then
                    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = A
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                            Impre(4, 1) = 0
                            Impre(4, 2) = 0
                            Impre(5, 1) = 0
                            Impre(5, 2) = 0
                            Impre(6, 1) = 0
                            Impre(6, 2) = 0
                            Impre(7, 1) = 0
                            Impre(7, 2) = 0
                            Impre(8, 1) = 0
                            Impre(8, 2) = 0
                            Impre(9, 1) = 0
                            Impre(9, 2) = 0
                            Impre(10, 1) = 0
                            Impre(10, 2) = 0
    
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
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
                               
                                            If rstLaudo!Saldo <> 0 Then
                                                If rstLaudo!Articulo = Articulo Then
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstLaudo!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstLaudo!Laudo
                                                            Impre(Xlugar, 2) = rstLaudo!Saldo
                                                            XCanti = XCanti - rstLaudo!Saldo
                                                        End If
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
                        
                            XParam = "'" + Articulo + "','" _
                                         + Articulo + "'"
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
                               
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Articulo = Articulo Then
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                                rstMovguia.Close
                            End If
                            
                        End If
                        
                        WArticulo1 = Left$(Articulo, 2)
                        WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
                        WCantidad = Cantidad
                        
                        ZLugar = 0
                        ZSuma = 0
                        
                        For ZCiclo = 1 To 10
                        
                            ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                            ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                            
                            If Impre(ZCiclo, 2) <> 0 Then
                                ZSuma = ZSuma + Impre(ZCiclo, 2)
                                    Else
                                If ZLugar = 0 Then
                                    ZLugar = ZCiclo
                                End If
                            End If
                            
                        Next ZCiclo
                        
                        If Val(CantidadII) > ZSuma Then
                            ZZCanti(ZLugar) = Str$(Val(CantidadII) - ZSuma)
                            ZZLote(ZLugar) = ""
                        End If
                        
                        
                        
                        
                        Linea = Linea + 1
                        WLinea = Str$(Linea)
                        
                        For ZCiclo = 1 To 10
                        
                            If Val(ZZCanti(ZCiclo)) <> 0 Then
                        
                                LineaII = LineaII + 1
                                WLIneaII = Str$(LineaII)
                            
                                WCantidadII = ZZCanti(ZCiclo)
                                WLoteII = ZZLote(ZCiclo)
                                
                                ZSql = ""
                                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                                ZSql = ZSql & "Hoja ,"
                                ZSql = ZSql & "Renglon ,"
                                ZSql = ZSql & "Fecha ,"
                                ZSql = ZSql & "Codigo1 ,"
                                ZSql = ZSql & "Codigo2 ,"
                                ZSql = ZSql & "Maquina ,"
                                ZSql = ZSql & "Articulo1 ,"
                                ZSql = ZSql & "Articulo2 ,"
                                ZSql = ZSql & "Cantidad ,"
                                ZSql = ZSql & "Lote ,"
                                ZSql = ZSql & "Teorico ,"
                                ZSql = ZSql & "Terminado ,"
                                ZSql = ZSql & "Equipo )"
                                ZSql = ZSql & "Values ("
                                ZSql = ZSql & "'" + WHoja + "',"
                                ZSql = ZSql & "'" + WLIneaII + "',"
                                ZSql = ZSql & "'" + WFecha + "',"
                                ZSql = ZSql & "'" + WCodigo1 + "',"
                                ZSql = ZSql & "'" + WCodigo2 + "',"
                                ZSql = ZSql & "'" + WMaquina + "',"
                                ZSql = ZSql & "'" + WArticulo1 + "',"
                                ZSql = ZSql & "'" + WArticulo2 + "',"
                                ZSql = ZSql & "'" + WCantidadII + "',"
                                ZSql = ZSql & "'" + WLoteII + "',"
                                ZSql = ZSql & "'" + WTeorico + "',"
                                ZSql = ZSql & "'" + Producto.Text + "',"
                                ZSql = ZSql & "'" + Equipo.Text + "')"
        
                                spImpreHojaII = ZSql
                                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                            
                            End If
                            
                        Next ZCiclo
                        
                End If

                If Tipo = "T" Then
                    
                        Erase Impre
                        Xlugar = 0
                        XCanti = Val(Cantidad)
                        
                        ZLugar = A
                        Impre(1, 1) = Val(ZLote(ZLugar, 1))
                        Impre(1, 2) = Val(ZLote(ZLugar, 2))
                        Impre(2, 1) = Val(ZLote(ZLugar, 3))
                        Impre(2, 2) = Val(ZLote(ZLugar, 4))
                        Impre(3, 1) = Val(ZLote(ZLugar, 5))
                        Impre(3, 2) = Val(ZLote(ZLugar, 6))
                        
                        If Impre(1, 1) = 0 Or Impre(1, 2) = 0 Then
                        
                            Impre(1, 1) = 0
                            Impre(1, 2) = 0
                            Impre(2, 1) = 0
                            Impre(2, 2) = 0
                            Impre(3, 1) = 0
                            Impre(3, 2) = 0
                            Impre(4, 1) = 0
                            Impre(4, 2) = 0
                            Impre(5, 1) = 0
                            Impre(5, 2) = 0
                            Impre(6, 1) = 0
                            Impre(6, 2) = 0
                            Impre(7, 1) = 0
                            Impre(7, 2) = 0
                            Impre(8, 1) = 0
                            Impre(8, 2) = 0
                            Impre(9, 1) = 0
                            Impre(9, 2) = 0
                            Impre(10, 1) = 0
                            Impre(10, 2) = 0
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
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
                            
                                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                                If rstHoja!Producto = Terminado Then
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstHoja!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                                            XCanti = XCanti - rstHoja!Saldo
                                                        End If
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
                                rstHoja.Close
                            End If
                        
                            XParam = "'" + Terminado + "','" _
                                         + Terminado + "'"
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
                            
                                            If rstMovguia!Saldo <> 0 Then
                                                If rstMovguia!Terminado = Terminado Then
                                                    If Xlugar < 10 And XCanti > 0 Then
                                                        Xlugar = Xlugar + 1
                                                        If rstMovguia!Saldo > XCanti Then
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = XCanti
                                                            XCanti = 0
                                                                Else
                                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                                            XCanti = XCanti - rstMovguia!Saldo
                                                        End If
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
                                rstMovguia.Close
                            End If
                        
                        End If

                        Linea = Linea + 1
                        WLinea = Str$(Linea)
                        
                        WArticulo1 = Left$(Terminado, 2)
                        WArticulo2 = Mid$(Terminado, 4, 5) + "-" + Right$(Terminado, 3)
                        
                        WCantidad = Cantidad
                        
                        For ZCiclo = 1 To 10
                            ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                            ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
                        Next ZCiclo
                        
                        For ZCiclo = 1 To 10
                        
                            If Val(ZZCanti(ZCiclo)) <> 0 Then
                        
                                LineaII = LineaII + 1
                                WLIneaII = Str$(LineaII)
                            
                                WCantidadII = ZZCanti(ZCiclo)
                                WLoteII = ZZLote(ZCiclo)
                        
                                ZSql = ""
                                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                                ZSql = ZSql & "Hoja ,"
                                ZSql = ZSql & "Renglon ,"
                                ZSql = ZSql & "Fecha ,"
                                ZSql = ZSql & "Codigo1 ,"
                                ZSql = ZSql & "Codigo2 ,"
                                ZSql = ZSql & "Maquina ,"
                                ZSql = ZSql & "Articulo1 ,"
                                ZSql = ZSql & "Articulo2 ,"
                                ZSql = ZSql & "Cantidad ,"
                                ZSql = ZSql & "Lote ,"
                                ZSql = ZSql & "Teorico ,"
                                ZSql = ZSql & "Terminado ,"
                                ZSql = ZSql & "Equipo )"
                                ZSql = ZSql & "Values ("
                                ZSql = ZSql & "'" + WHoja + "',"
                                ZSql = ZSql & "'" + WLIneaII + "',"
                                ZSql = ZSql & "'" + WFecha + "',"
                                ZSql = ZSql & "'" + WCodigo1 + "',"
                                ZSql = ZSql & "'" + WCodigo2 + "',"
                                ZSql = ZSql & "'" + WMaquina + "',"
                                ZSql = ZSql & "'" + WArticulo1 + "',"
                                ZSql = ZSql & "'" + WArticulo2 + "',"
                                ZSql = ZSql & "'" + WCantidadII + "',"
                                ZSql = ZSql & "'" + WLoteII + "',"
                                ZSql = ZSql & "'" + WTeorico + "',"
                                ZSql = ZSql & "'" + Producto.Text + "',"
                                ZSql = ZSql & "'" + Equipo.Text + "')"
            
                                spImpreHojaII = ZSql
                                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                            
                            End If
                            
                        Next ZCiclo
                        
                End If
            
            Next A
            
            XLinea = LineaII
            For Ciclo = XLinea To 24
            
                LineaII = LineaII + 1
                WLIneaII = Str$(LineaII)
                        
                WCantidadII = ""
                WLoteII = ""
                WArticulo1 = ""
                WArticulo2 = ""
                                                   
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                ZSql = ZSql & "Hoja ,"
                ZSql = ZSql & "Renglon ,"
                ZSql = ZSql & "Fecha ,"
                ZSql = ZSql & "Codigo1 ,"
                ZSql = ZSql & "Codigo2 ,"
                ZSql = ZSql & "Maquina ,"
                ZSql = ZSql & "Articulo1 ,"
                ZSql = ZSql & "Articulo2 ,"
                ZSql = ZSql & "Cantidad ,"
                ZSql = ZSql & "Lote ,"
                ZSql = ZSql & "Teorico ,"
                ZSql = ZSql & "Terminado ,"
                ZSql = ZSql & "Equipo )"
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WHoja + "',"
                ZSql = ZSql & "'" + WLIneaII + "',"
                ZSql = ZSql & "'" + WFecha + "',"
                ZSql = ZSql & "'" + WCodigo1 + "',"
                ZSql = ZSql & "'" + WCodigo2 + "',"
                ZSql = ZSql & "'" + WMaquina + "',"
                ZSql = ZSql & "'" + WArticulo1 + "',"
                ZSql = ZSql & "'" + WArticulo2 + "',"
                ZSql = ZSql & "'" + WCantidadII + "',"
                ZSql = ZSql & "'" + WLoteII + "',"
                ZSql = ZSql & "'" + WTeorico + "',"
                ZSql = ZSql & "'" + Producto.Text + "',"
                ZSql = ZSql & "'" + Equipo.Text + "')"
        
                spImpreHojaII = ZSql
                Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)

            Next Ciclo
            
            Listado.WindowTitle = "Impresion de Hoja de Produccion"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
   
            Listado.Destination = 1
            Rem Listado.Destination = 0
            Listado.CopiesToPrinter = 2
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Select Case XTipoPro
                Case "CO"
                    Listado.ReportFileName = "ImpreHojaNuevoIIA4.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
                        
                    Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                                + "From " _
                                + DSQ + ".dbo.ImpreHojaII ImpreHojaII " _
                                + "Where " _
                                + "ImpreHojaII.Hoja >= 0 AND " _
                                + "ImpreHojaII.Hoja <= 999999"
            
                Case Else
                    Listado.ReportFileName = "ImpreHojaNuevoII.rpt"
                    Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
                        
                    Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                                + "From " _
                                + DSQ + ".dbo.ImpreHojaII ImpreHojaII " _
                                + "Where " _
                                + "ImpreHojaII.Hoja >= 0 AND " _
                                + "ImpreHojaII.Hoja <= 999999"
            End Select
            Listado.Connect = Connect()
            Listado.Action = 1
            Listado.CopiesToPrinter = 1
        
        Case Else
            
    End Select
    
    Call Limpia_Click
    Hoja.SetFocus
        
    Exit Sub

WError:

    Resume Next
    
End Sub
    
Sub Etiqueta()
    PCliente = Cliente.Text
    PTipo = 1
    PrgConsCcte.Show
End Sub

Sub Calcula_stock()

    Muestra.Clear
    Muestra.Row = 0
 
    Muestra.Col = 1
    Muestra.Text = "Tipo"
    
    Muestra.Col = 2
    Muestra.Text = "Partida"
    
    Muestra.Col = 3
    Muestra.Text = "Stock"
    
    Muestra1.Clear
    Muestra1.Row = 0
 
    Muestra1.Col = 1
    Muestra1.Text = "Cliente"
    
    Muestra1.Col = 3
    Muestra1.Text = "Fecha"
    
    Muestra1.Col = 2
    Muestra1.Text = "Cantidad"
    
    Producto.Text = UCase(Producto.Text)
    XProducto = Producto.Text
    Renglon = 0
    XStock = 0
    XPedido = 0
    
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
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "PT"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        rstHoja.Close
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "PT"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        
        rstMovguia.Close
        
    End If
    
    
    Rem lee NK
    
    XProducto = "NK" + Mid$(XProducto, 3, 10)
    
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
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "NK"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        rstHoja.Close
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "NK"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        rstMovguia.Close
    End If
    
    
    Rem lee re
    
    XProducto = "RE" + Mid$(XProducto, 3, 10)
    
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
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "RE"
                
                        Muestra.Col = 2
                        Muestra.Text = XHoja
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        rstHoja.Close
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    XLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    XSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(XSaldo)
                    
                    If XSaldo <> 0 Then
                    
                        Renglon = Renglon + 1
                        Muestra.Row = Renglon
                        
                        Muestra.Col = 1
                        Muestra.Text = "RE"
                
                        Muestra.Col = 2
                        Muestra.Text = XLote
                        
                        Muestra.Col = 3
                        Muestra.Text = Pusing("###,###.##", Str$(XSaldo))
                        
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
        rstMovguia.Close
    End If
    
    spPedido = "ListaPedidoTerminado " + "'" + Producto.Text + "'"
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
                
                If XPed <> 0 And Left$(rstPedido!Cliente, 1) = "X" Then
                    
                    Renglon = Renglon + 1
                    Muestra1.Row = Renglon
                    
                    Muestra1.Col = 1
                    Muestra1.Text = rstPedido!Cliente
                
                    Muestra1.Col = 3
                    Muestra1.Text = rstPedido!Fecha
                        
                    Muestra1.Col = 2
                    Muestra1.Text = Pusing("###,###.##", Str$(XPed))
                        
                    XPedido = XPedido + XPed
                        
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        rstPedido.Close
    End If
    
    Stock.Text = Pusing("###,###.##", Str$(XStock))
    Pedido.Text = Pusing("###,###.##", Str$(XPedido))

End Sub


Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True

    
    WSaldo1 = 0
    WSaldo2 = 0
    WSaldo3 = 0
    
    If Val(WLote1.Text) <> 0 Then
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
            
                XParam = "'" + WLote1.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote1.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote1.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
        
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.###", WCanti1.Text)
            WControl1.Locked = False
            WControl1.Text = "X"
            WControl1.Locked = True
        End If
        
    End If
    
    If Val(WLote2.Text) <> 0 Then
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote2.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote2.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
            
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.###", WCanti2.Text)
            WControl2.Locked = False
            WControl2.Text = "X"
            WControl2.Locked = True
        End If
        
    End If
    
    
    If Val(WLote3.Text) <> 0 Then
        If WWTipo = "M" Then
        
            WEntra = "N"
        
            WControla = 0
            spArticulo = "ConsultaArticulo " + "'" + WWArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WControla = IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla)
                rstArticulo.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                            + WWArticulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WEntra = "S"
                    rstLaudo.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWArticulo + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
            End If
            
            If WEntra <> "S" Then
                m$ = WWArticulo + " Articulo inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
                Else
        
            WEntra = "N"
            
            WControla = 0
            spTerminado = "ConsultaTerminado " + "'" + WWTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                rstTerminado.Close
            End If
            
            If WControla = 0 Then
                XParam = "'" + WLote3.Text + "','" _
                        + WWTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    WEntra = "S"
                    rstHoja.Close
                End If
                
                If WEntra = "N" Then
                    XParam = "'" + WWTerminado + "','" _
                            + WLote3.Text + "'"
                    spMovguia = "ListaMovguiaLote1 " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        WEntra = "S"
                        rstMovguia.Close
                    End If
                End If
                
                    Else
                    
                WEntra = "S"
                
            End If
                
            If WEntra <> "S" Then
                m$ = WWTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
        End If
        
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.###", WCanti3.Text)
            WControl3.Locked = False
            WControl3.Text = "X"
            WControl3.Locked = True
        End If
        
    End If
    
    If Val(WLote1.Text) <> 0 And WControl1.Text = "X" Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If Val(WLote2.Text) <> 0 And WControl2.Text = "X" Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Val(WLote3.Text) <> 0 And WControl3.Text = "X" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    
    If Suma = Val(WWCantidad) Then
        WEstado = "S"
    End If
    
    If WControla <> 0 Then
        WEstado = "S"
    End If
    
    
End Sub

Private Sub WLote1_DblClick()
    ZProceso = 1
    If WWTipo = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub WLote2_DblClick()
    ZProceso = 2
    If WWTipo = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub WLote3_DblClick()
    ZProceso = 3
    If WWTipo = "M" Then
        Call ficha_Mp
            Else
        Call ficha_Pt
    End If
End Sub

Private Sub ficha_Mp()

    Call Limpia_Vector
    
    XRenglon = 0
    XParam = "'" + WWArticulo + "','" _
                 + WWArticulo + "'"
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
                    
                    If rstLaudo!Articulo = WWArticulo Then
                
                        ZArticulo = rstLaudo!Articulo
                        ZCantidad = rstLaudo!Liberada
                        ZFecha = rstLaudo!Fecha
                        ZLaudo = rstLaudo!Laudo
                        ZOrden = rstLaudo!Orden
                        Zdevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        ZRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        ZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        ZLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        Call Redondeo(ZSaldo)
                        
                        If ZLiberada <> 0 And ZSaldo <> 0 Then
                        
                            XRenglon = XRenglon + 1
                            WVector1.Row = XRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = "Laudo"
                        
                            WVector1.Col = 2
                            WVector1.Text = ZLaudo
                                               
                            WVector1.Col = 3
                            WVector1.Text = ZFecha
                        
                            WVector1.Col = 4
                            WVector1.Text = ZOrden
                        
                            WVector1.Col = 5
                            WVector1.Text = ZCantidad
                
                            WVector1.Col = 6
                            WVector1.Text = ZSaldo
                
                            WVector1.Col = 7
                            WVector1.Text = ZLaudo
                            
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
    
    XParam = "'" + WWArticulo + "','" _
                + WWArticulo + "'"
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
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WWArticulo Then
                    
                        ZArticulo = rstMovguia!Articulo
                        ZCantidad = rstMovguia!Cantidad
                        ZFecha = rstMovguia!Fecha
                        ZCodigo = rstMovguia!Codigo
                        ZMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        ZTipomov = rstMovguia!Tipomov
                        ZPartida = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        ZFecha = rstMovguia!Fecha
                        If Val(ZCodigo) > 900000 Then
                            ZTipo = "Prestamo"
                            ZCodigo = ZCodigo - 900000
                                Else
                            ZTipo = "Guia In"
                        End If
                        ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZSaldo)
                                
                        If rstMovguia!Movi = "E" And ZSaldo <> 0 Then
                            
                            XRenglon = XRenglon + 1
                            WVector1.Row = XRenglon
                
                            WVector1.Col = 1
                            WVector1.Text = ZTipo
                        
                            WVector1.Col = 2
                            WVector1.Text = ZCodigo
                                               
                            WVector1.Col = 3
                            WVector1.Text = ZFecha
                        
                            WVector1.Col = 4
                            WVector1.Text = ""
                        
                            WVector1.Col = 5
                            WVector1.Text = ZCantidad
                
                            WVector1.Col = 6
                            WVector1.Text = ZSaldo
                
                            WVector1.Col = 7
                            WVector1.Text = ZPartida
                            
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
        rstMovguia.Close
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
End Sub

Private Sub ficha_Pt()

    Call Limpia_Vector
    
    XRenglon = 0
    
    XParam = "'" + WWTerminado + "','" _
                 + WWTerminado + "'"
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
                Rem And rstHoja!Real <> 0 Then
                 
                    ZProducto = rstHoja!Producto
                    ZCantidad = rstHoja!Real
                    ZFecha = rstHoja!Fecha
                    ZHoja = rstHoja!Hoja
                    ZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZSaldo)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector1.Row = XRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = "Hoja"
                        
                        WVector1.Col = 2
                        WVector1.Text = ZHoja
                                               
                        WVector1.Col = 3
                        WVector1.Text = ZFecha
                        
                        WVector1.Col = 4
                        WVector1.Text = ""
                        
                        WVector1.Col = 5
                        WVector1.Text = ZCantidad
                
                        WVector1.Col = 6
                        WVector1.Text = ZSaldo
                
                        WVector1.Col = 7
                        WVector1.Text = ZHoja
                    
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
        rstHoja.Close
    End If
    
    
    
    XParam = "'" + WWTerminado + "','" _
                 + WWTerminado + "'"
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    ZTerminado = rstMovguia!Terminado
                    ZCantidad = rstMovguia!Cantidad
                    ZFecha = rstMovguia!Fecha
                    ZCodigo = rstMovguia!Codigo
                    ZMovi = rstMovguia!Movi
                    ZDestino = rstMovguia!Destino
                    ZTipomov = rstMovguia!Tipomov
                    WWLote = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                    ZPartida = IIf(IsNull(rstMovguia!Partida), "", rstMovguia!Partida)
                    ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZSaldo)
                    If Val(ZCodigo) > 900000 Then
                        ZTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        ZTipo = "Guia In"
                    End If
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector1.Row = XRenglon
                
                        WVector1.Col = 1
                        WVector1.Text = ZTipo
                        
                        WVector1.Col = 2
                        WVector1.Text = ZCodigo
                                               
                        WVector1.Col = 3
                        WVector1.Text = ZFecha
                        
                        WVector1.Col = 4
                        WVector1.Text = ""
                        
                        WVector1.Col = 5
                        WVector1.Text = ZCantidad
                
                        WVector1.Col = 6
                        WVector1.Text = ZSaldo
                
                        WVector1.Col = 7
                        WVector1.Text = WWLote
                        
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
        
        rstMovguia.Close
        
    End If
    
    
    
    XParam = "'" + WWTerminado + "','" _
                 + WWTerminado + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntdev!Marca = "X" Then
                
                        Else
                
                ZTerminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                WWLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector1.Row = XRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = "Dev"
                        
                    WVector1.Col = 2
                    WVector1.Text = ZCodigo
                                               
                    WVector1.Col = 3
                    WVector1.Text = ZFecha
                        
                    WVector1.Col = 4
                    WVector1.Text = ""
                        
                    WVector1.Col = 5
                    WVector1.Text = ZCantidad
                
                    WVector1.Col = 6
                    WVector1.Text = ZSaldo
                
                    WVector1.Col = 7
                    WVector1.Text = WWLote
                        
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        
        rstEntdev.Close
        
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector1.TopRow = 1
    
End Sub


Private Sub Limpia_Vector()

    WVector1.Height = 4095
    WVector1.Left = 120
    WVector1.Top = 1200
    WVector1.Width = 10000

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 5001
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector1.Text = "Orden"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Saldo"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
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
    
    WVector1.Visible = True
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Click()
    busquedalote = WVector1.TextMatrix(WVector1.Row, 7)
    WVector1.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    WTitulo(5).Visible = False
    WTitulo(6).Visible = False
    WTitulo(7).Visible = False
    Select Case ZProceso
        Case 1
            WLote1.Text = busquedalote
            WCanti1.SetFocus
        Case 2
            WLote2.Text = busquedalote
            WCanti2.SetFocus
        Case 3
            WLote3.Text = busquedalote
            WCanti3.SetFocus
        Case Else
    End Select
        
End Sub



Private Sub Calcula_Desvio()
                
    ZZDesvio = 0
    
    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + ZZArticulo + "','" _
                 + ZZArticulo + "'"
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
                    
                    If rstLaudo!Articulo = ZZArticulo Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!Liberada
                        WLaudo = rstLaudo!Laudo
                        ZZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        Call Redondeo(ZZSaldo)
                        
                        If WLaudo >= 190000 And WLaudo <= 194999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 990000 And WLaudo <= 994999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 290000 And WLaudo <= 294999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 390000 And WLaudo <= 394999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 490000 And WLaudo <= 494999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 590000 And WLaudo <= 594999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 690000 And WLaudo <= 694999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 790000 And WLaudo <= 794999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 890000 And WLaudo <= 894999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
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
    
    XParam = "'" + ZZArticulo + "','" _
                 + ZZArticulo + "'"
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
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = ZZArticulo Then
                    
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZZSaldo)
                        WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        
                        If WLaudo >= 190000 And WLaudo <= 194999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 990000 And WLaudo <= 994999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 290000 And WLaudo <= 294999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 390000 And WLaudo <= 394999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 490000 And WLaudo <= 494999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 590000 And WLaudo <= 594999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 690000 And WLaudo <= 694999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 790000 And WLaudo <= 794999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
                        End If
                        If WLaudo >= 890000 And WLaudo <= 894999 Then
                            ZZDesvio = ZZDesvio + ZZSaldo
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
        rstMovguia.Close
    End If
    
End Sub


Private Sub Calcula_Bloqueo()
                
    ZZBloqueo = 0
    Rem dada
    
    If ZZTipoBloqueo = "M" Then
    
        Rem PROCESA LOS LAUDOS
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            With rstLaudo
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstLaudo.Close
        End If
    
        Rem PROCESA LOS GUIAS DE TRASLADO INTERN
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstMovguia.Close
        End If
        
            Else
    
        Rem PROCESA LAS HOJAS
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Producto = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Renglon = 1"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
        
            With rstHoja
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstHoja.Close
        End If
    
        Rem PROCESA LOS GUIAS DE TRASLADO INTERN
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Articulo = " + "'" + ZZCodigoBloqueo + "'"
        ZSql = ZSql + " and Saldo <> 0"
        ZSql = ZSql + " and Estado = " + "'" + "N" + "'"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
        
            With rstMovguia
        
                .MoveFirst
                
                If .NoMatch = False Then
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    ZZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                    Call Redondeo(ZZSaldo)
                    
                    ZZBloqueo = ZZBloqueo + ZZSaldo
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
                End If
            End With
            rstMovguia.Close
        End If
            
    End If
    
End Sub

Private Sub Calcula_Stock_Vencido()

    If WWArticulo = "AA-100-100" Then
        Exit Sub
    End If
    
    Erase ZArti
    ZLugar = 0
    WWVencido = 0
    
    If WWArticulo = "AA-100-100" Then
        Exit Sub
    End If
    
    Rem PROCESA LOS LAUDOS
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Articulo = " + "'" + WWArticulo + "'"
    ZSql = ZSql + " and Saldo <> 0"
    spLaudo = ZSql
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
                    
                    XXLaudo = rstLaudo!Laudo
                    XXSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    Call Redondeo(XXSaldo)
                    XXArticulo = rstLaudo!Articulo
                    XXCantidad = rstLaudo!Liberada
                    XXFecha = rstLaudo!Fecha
                    XXClave = rstLaudo!Clave
                    
                    If XXSaldo <> 0 Then
                        ZLugar = ZLugar + 1
                        ZArti(ZLugar, 1) = XXLaudo
                        ZArti(ZLugar, 2) = XXArticulo
                        ZArti(ZLugar, 3) = Str$(XXCantidad)
                        ZArti(ZLugar, 4) = Str$(XXSaldo)
                        ZArti(ZLugar, 5) = XXFecha
                        ZArti(ZLugar, 6) = XXClave
                        ZArti(ZLugar, 7) = "L"
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
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Articulo = " + "'" + WWArticulo + "'"
    ZSql = ZSql + " and Saldo <> 0"
    spMovguia = ZSql
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 And rstMovguia!Codigo > 900000 Then
                
                        Else
                        
                    If rstMovguia!Tipo = "M" Then
                    
                        XXLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        XXArticulo = rstMovguia!Articulo
                        XXCantidad = rstMovguia!Cantidad
                        XXSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(XXSaldo)
                        XXFecha = rstMovguia!Fecha
                        XXClave = rstMovguia!Clave
                        
                        If XXSaldo <> 0 Then
                            ZLugar = ZLugar + 1
                            ZArti(ZLugar, 1) = XXLaudo
                            ZArti(ZLugar, 2) = XXArticulo
                            ZArti(ZLugar, 3) = Str$(XXCantidad)
                            ZArti(ZLugar, 4) = Str$(XXSaldo)
                            ZArti(ZLugar, 5) = XXFecha
                            ZArti(ZLugar, 6) = XXClave
                            ZArti(ZLugar, 7) = "G"
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
        rstMovguia.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZVto = ""
        
        ZLaudo = ZArti(Ciclo, 1)
        ZArticulo = ZArti(Ciclo, 2)
        ZCantidad = Val(ZArti(Ciclo, 3))
        ZSaldo = Val(ZArti(Ciclo, 4))
        ZFecha = ZArti(Ciclo, 5)
        ZClave = ZArti(Ciclo, 6)
        ZTipo = ZArti(Ciclo, 7)
        ZMarcaVencida = ""
        
        XEmpresa = Wempresa
    
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0002"
        Empe(2, 2) = "Empresa02"
        Empe(3, 1) = "0003"
        Empe(3, 2) = "Empresa03"
        Empe(4, 1) = "0004"
        Empe(4, 2) = "Empresa04"
        Empe(5, 1) = "0005"
        Empe(5, 2) = "Empresa05"
        Empe(6, 1) = "0006"
        Empe(6, 2) = "Empresa06"
        Empe(7, 1) = "0007"
        Empe(7, 2) = "Empresa07"
        Empe(8, 1) = "0008"
        Empe(8, 2) = "Empresa08"
        Empe(9, 1) = "0009"
        Empe(9, 2) = "Empresa09"
        Empe(10, 1) = "0010"
        Empe(10, 2) = "Empresa10"
        Empe(11, 1) = "0011"
        Empe(11, 2) = "Empresa11"
        
        XHasta = 11
                
        For Ciclo2 = 1 To XHasta
    
            Wempresa = Empe(Ciclo2, 1)
            txtOdbc = Empe(Ciclo2, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo = " + "'" + ZLaudo + "'"
            ZSql = ZSql + " and Articulo = " + "'" + ZArticulo + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                ZFecha = rstLaudo!Fecha
                ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                rstLaudo.Close
                Exit For
            End If
            
        Next Ciclo2
            
        Call Conecta_Empresa
    
        ZVto = ""
        ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        If ZFechaVto <> "" And ZFechaVto <> "  /  /    " And ZFechaVto <> "00/00/0000" Then
            Call Valida_fecha(ZFechaVto, Auxi)
            If Auxi = "S" Then
                ZVto = ZFechaVto
            End If
        End If
            
        If ZVto = "" Then
            
            ZMeses = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZMeses = rstArticulo!Meses
                rstArticulo.Close
            End If
            
            WMes = Val(Mid$(ZFecha, 4, 2))
            WAno = Val(Right$(ZFecha, 4))
            For ZCiclo = 1 To ZMeses
                WMes = WMes + 1
                If WMes > 12 Then
                    WAno = WAno + 1
                    WMes = 1
                End If
            Next ZCiclo
            
            XMes = Str$(WMes)
            XAno = Str$(WAno)
            Call Ceros(XMes, 2)
            Call Ceros(XAno, 4)
            If Val(Left$(ZFecha, 2)) <= 30 Then
                If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
                End If
                    Else
                If Val(XMes) = 2 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = "30/" + XMes + "/" + XAno
                End If
            End If
            
        End If
        
        If ZFecha <> "" Then
        
            Do
                Call Valida_fecha(ZVto, Auxi)
                If Auxi = "S" Then
                    Exit Do
                        Else
                    XFec1 = ZVto
                    SumaDia = 1
                    Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                    ZVto = XFec2
                End If
            Loop
        
            Rem WFechaActual = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            Rem WFechaVto = Right$(ZVto, 4) + Mid$(ZVto, 4, 2) + Left$(ZVto, 2)
            ZZComparaI = Fecha.Text
            If Left$(ZVto, 2) > "28" Then
                ZZComparaII = "28" + Mid$(ZVto, 3, 8)
                    Else
                ZZComparaII = ZVto
            End If
        
            ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
        
            If Val(ZDias) < 0 Then
                ZMarcaVencida = "S"
                WWVencido = WWVencido + ZSaldo
            End If
        
        End If
        
        If ZTipo = "L" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Laudo SET "
            ZSql = ZSql + " MarcaVencida = " + "'" + ZMarcaVencida + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "UPDATE Guia SET "
            ZSql = ZSql + " MarcaVencida = " + "'" + ZMarcaVencida + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Next Ciclo
End Sub
                
Private Sub ImpreHojaAlmacero()

    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WCodigo1 = Left$(Producto.Text, 2)
    WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
    WMaquina = ""
    WTeorico = Teorico.Text
            
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZMetodo = IIf(IsNull(rstTerminado!Metodo), "0", rstTerminado!Metodo)
        ZEfluentes = IIf(IsNull(rstTerminado!Efluentes), "0", rstTerminado!Efluentes)
        ZVersionI = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
        ZVersionII = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
        ZVersionIII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
        rstTerminado.Close
    End If
            
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Efluentes"
    ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + ZEfluentes + "'"
    spEfluentes = ZSql
    Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
    If rstEfluentes.RecordCount > 0 Then
        ZDesEfluentesI = rstEfluentes!Descripcion
        ZDesEfluentesII = rstEfluentes!DescripcionII
        rstEfluentes.Close
    End If

            
    Sql1 = "DELETE ImpreHojaII"
    spImpreHojaII = Sql1
    Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                
    Linea = 0
    LineaII = 0
    ZZLugar = 0
        
    For A = 1 To 40
        
        ZZLugar = ZZLugar + 1
            
        Tipo = UCase(Grilla.TextMatrix(A, 1))
        Terminado = UCase(Grilla.TextMatrix(A, 2))
        Articulo = UCase(Grilla.TextMatrix(A, 3))
        Cantidad = Grilla.TextMatrix(A, 5)
             
        If Tipo = "M" Then
                
            Erase Impre
            Xlugar = 0
            XCanti = Val(Cantidad)
                    
            ZLugar = A
            
            Erase Impre
                        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Laudo.Laudo"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                        
                With rstLaudo
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                           
                            WMarcaVencida = IIf(IsNull(rstLaudo!MarcaVencida), "", rstLaudo!MarcaVencida)
                            QSaldo = rstLaudo!Saldo
                            Call Redondeo(QSaldo)
                            If QSaldo <> 0 And Trim(WMarcaVencida) = "" Then
                                If rstLaudo!Articulo = Articulo Then
                                            
                                    WLaudo = rstLaudo!Laudo
                                    ZEntra = "S"
                                                
                                    If WLaudo >= 190000 And WLaudo <= 194999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 990000 And WLaudo <= 994999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 290000 And WLaudo <= 294999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 390000 And WLaudo <= 394999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 490000 And WLaudo <= 494999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 590000 And WLaudo <= 594999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 690000 And WLaudo <= 694999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 790000 And WLaudo <= 794999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 890000 And WLaudo <= 894999 Then
                                        ZEntra = "N"
                                    End If
                                                
                                    If ZEntra = "S" Then
                                        If Xlugar < 10 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstLaudo!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstLaudo!Laudo
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                    Else
                                                Impre(Xlugar, 1) = rstLaudo!Laudo
                                                Impre(Xlugar, 2) = rstLaudo!Saldo
                                                XCanti = XCanti - rstLaudo!Saldo
                                            End If
                                        End If
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
                    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                           
                With rstMovguia
                           
                    .MoveFirst
                           
                    If .NoMatch = False Then
                        Do
                           
                            If .EOF = True Then
                                Exit Do
                            End If
                        
                            WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                            QSaldo = rstMovguia!Saldo
                            Call Redondeo(QSaldo)
                            If QSaldo <> 0 And Trim(WMarcaVencida) = "" Then
                                If rstMovguia!Articulo = Articulo Then
                                            
                                    WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                                    ZEntra = "S"
                                                    
                                    If WLaudo >= 190000 And WLaudo <= 194999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 990000 And WLaudo <= 994999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 290000 And WLaudo <= 294999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 390000 And WLaudo <= 394999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 490000 And WLaudo <= 494999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 590000 And WLaudo <= 594999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 690000 And WLaudo <= 694999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 790000 And WLaudo <= 794999 Then
                                        ZEntra = "N"
                                    End If
                                    If WLaudo >= 890000 And WLaudo <= 894999 Then
                                        ZEntra = "N"
                                    End If
                                            
                                    If ZEntra = "S" Then
                                            
                                        If Xlugar < 10 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstMovguia!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                    Else
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = rstMovguia!Saldo
                                                XCanti = XCanti - rstMovguia!Saldo
                                            End If
                                        End If
                                                    
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
                rstMovguia.Close
            End If
                    
            WArticulo1 = Left$(Articulo, 2)
            WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
            WCantidad = Cantidad
                    
            For ZCiclo = 1 To 10
                ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
            Next ZCiclo
                    
            Linea = Linea + 1
            WLinea = Str$(Linea)
                    
            For ZCiclo = 1 To 10
                    
                If Val(ZZCanti(ZCiclo)) <> 0 Then
                    
                    LineaII = LineaII + 1
                    WLIneaII = Str$(LineaII)
                        
                    WCantidadII = ZZCanti(ZCiclo)
                    WLoteII = ZZLote(ZCiclo)
                            
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                    ZSql = ZSql & "Hoja ,"
                    ZSql = ZSql & "Renglon ,"
                    ZSql = ZSql & "Fecha ,"
                    ZSql = ZSql & "Codigo1 ,"
                    ZSql = ZSql & "Codigo2 ,"
                    ZSql = ZSql & "Maquina ,"
                    ZSql = ZSql & "Articulo1 ,"
                    ZSql = ZSql & "Articulo2 ,"
                    ZSql = ZSql & "Cantidad ,"
                    ZSql = ZSql & "Lote ,"
                    ZSql = ZSql & "Teorico ,"
                    ZSql = ZSql & "Terminado ,"
                    ZSql = ZSql & "Equipo )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + WHoja + "',"
                    ZSql = ZSql & "'" + WLIneaII + "',"
                    ZSql = ZSql & "'" + WFecha + "',"
                    ZSql = ZSql & "'" + WCodigo1 + "',"
                    ZSql = ZSql & "'" + WCodigo2 + "',"
                    ZSql = ZSql & "'" + WMaquina + "',"
                    ZSql = ZSql & "'" + WArticulo1 + "',"
                    ZSql = ZSql & "'" + WArticulo2 + "',"
                    ZSql = ZSql & "'" + WCantidadII + "',"
                    ZSql = ZSql & "'" + WLoteII + "',"
                    ZSql = ZSql & "'" + WTeorico + "',"
                    ZSql = ZSql & "'" + Producto.Text + "',"
                    ZSql = ZSql & "'" + Equipo.Text + "')"
    
                    spImpreHojaII = ZSql
                    Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                        
            Next ZCiclo
                    
        End If
        
        If Tipo = "T" Then
                
            Erase Impre
            Xlugar = 0
            XCanti = Val(Cantidad)
                    
            ZLugar = A
        
            Erase Impre
            
            XParam = "'" + Terminado + "','" _
                         + Terminado + "'"
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
                        
                            If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                If rstHoja!Producto = Terminado Then
                                    If Xlugar < 10 And XCanti > 0 Then
                                        Xlugar = Xlugar + 1
                                        If rstHoja!Saldo > XCanti Then
                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                            Impre(Xlugar, 2) = XCanti
                                            XCanti = 0
                                                Else
                                            Impre(Xlugar, 1) = rstHoja!Hoja
                                            Impre(Xlugar, 2) = rstHoja!Saldo
                                            XCanti = XCanti - rstHoja!Saldo
                                        End If
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
                rstHoja.Close
            End If
                    
            XParam = "'" + Terminado + "','" _
                         + Terminado + "'"
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
                        
                            If rstMovguia!Saldo <> 0 Then
                                If rstMovguia!Terminado = Terminado Then
                                    If Xlugar < 10 And XCanti > 0 Then
                                        Xlugar = Xlugar + 1
                                        If rstMovguia!Saldo > XCanti Then
                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                            Impre(Xlugar, 2) = XCanti
                                            XCanti = 0
                                                Else
                                            Impre(Xlugar, 1) = rstMovguia!Lote
                                            Impre(Xlugar, 2) = rstMovguia!Saldo
                                            XCanti = XCanti - rstMovguia!Saldo
                                        End If
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
                rstMovguia.Close
            End If
                    
            Linea = Linea + 1
            WLinea = Str$(Linea)
                
            WArticulo1 = Left$(Terminado, 2)
            WArticulo2 = Mid$(Terminado, 4, 5) + "-" + Right$(Terminado, 3)
            WCantidad = Cantidad
                    
            For ZCiclo = 1 To 10
                ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
            Next ZCiclo
                    
            For ZCiclo = 1 To 10
                    
                If Val(ZZCanti(ZCiclo)) <> 0 Then
            
                    LineaII = LineaII + 1
                    WLIneaII = Str$(LineaII)
                        
                    WCantidadII = ZZCanti(ZCiclo)
                    WLoteII = ZZLote(ZCiclo)
                    
                    ZSql = ""
                    ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                    ZSql = ZSql & "Hoja ,"
                    ZSql = ZSql & "Renglon ,"
                    ZSql = ZSql & "Fecha ,"
                    ZSql = ZSql & "Codigo1 ,"
                    ZSql = ZSql & "Codigo2 ,"
                    ZSql = ZSql & "Maquina ,"
                    ZSql = ZSql & "Articulo1 ,"
                    ZSql = ZSql & "Articulo2 ,"
                    ZSql = ZSql & "Cantidad ,"
                    ZSql = ZSql & "Lote ,"
                    ZSql = ZSql & "Teorico ,"
                    ZSql = ZSql & "Terminado ,"
                    ZSql = ZSql & "Equipo )"
                    ZSql = ZSql & "Values ("
                    ZSql = ZSql & "'" + WHoja + "',"
                    ZSql = ZSql & "'" + WLIneaII + "',"
                    ZSql = ZSql & "'" + WFecha + "',"
                    ZSql = ZSql & "'" + WCodigo1 + "',"
                    ZSql = ZSql & "'" + WCodigo2 + "',"
                    ZSql = ZSql & "'" + WMaquina + "',"
                    ZSql = ZSql & "'" + WArticulo1 + "',"
                    ZSql = ZSql & "'" + WArticulo2 + "',"
                    ZSql = ZSql & "'" + WCantidadII + "',"
                    ZSql = ZSql & "'" + WLoteII + "',"
                    ZSql = ZSql & "'" + WTeorico + "',"
                    ZSql = ZSql & "'" + Producto.Text + "',"
                    ZSql = ZSql & "'" + Equipo.Text + "')"
            
                    spImpreHojaII = ZSql
                    Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                        
            Next ZCiclo
                    
        End If
            
    Next A
    
    
    Listado.WindowTitle = "Impresion de Hoja de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Listado.ReportFileName = "ImpreHojaNuevoII.rpt"
    Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Renglon, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                    + "From " _
                    + DSQ + ".dbo.ImpreHojaII ImprehojaII " _
                    + "Where ImpreHojaII.Hoja >= 0 AND ImpreHojaII.Hoja <= 999999"
                    
    Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
    Listado.SelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"

    Listado.Connect = Connect()
    Listado.Action = 1

End Sub




                
Private Sub ImpreHojaDesvio()

    ZImpre = "N"

    WHoja = Hoja.Text
    WFecha = Fecha.Text
    WCodigo1 = Left$(Producto.Text, 2)
    WCodigo2 = Mid$(Producto.Text, 4, 5) + "/" + Right$(Producto.Text, 3)
    WMaquina = ""
    WTeorico = Teorico.Text
            
    spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZMetodo = IIf(IsNull(rstTerminado!Metodo), "0", rstTerminado!Metodo)
        ZEfluentes = IIf(IsNull(rstTerminado!Efluentes), "0", rstTerminado!Efluentes)
        ZVersionI = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
        ZVersionII = IIf(IsNull(rstTerminado!VersionI), "0", rstTerminado!VersionI)
        ZVersionIII = IIf(IsNull(rstTerminado!VersionII), "0", rstTerminado!VersionII)
        rstTerminado.Close
    End If
            
    ZSql = ""
    ZSql = ZSql & "Select *"
    ZSql = ZSql & " FROM Efluentes"
    ZSql = ZSql & " Where Efluentes.Codigo = " + "'" + ZEfluentes + "'"
    spEfluentes = ZSql
    Set rstEfluentes = db.OpenRecordset(spEfluentes, dbOpenSnapshot, dbSQLPassThrough)
    If rstEfluentes.RecordCount > 0 Then
        ZDesEfluentesI = rstEfluentes!Descripcion
        ZDesEfluentesII = rstEfluentes!DescripcionII
        rstEfluentes.Close
    End If

            
    Sql1 = "DELETE ImpreHojaII"
    spImpreHojaII = Sql1
    Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                
    Linea = 0
    LineaII = 0
    ZZLugar = 0
        
    For A = 1 To 40
        
        ZZLugar = ZZLugar + 1

        Tipo = UCase(Grilla.TextMatrix(A, 1))
        Terminado = UCase(Grilla.TextMatrix(A, 2))
        Articulo = UCase(Grilla.TextMatrix(A, 3))
        Cantidad = Grilla.TextMatrix(A, 5)
             
        If Tipo = "M" Then
                
            Erase Impre
            Xlugar = 0
            XCanti = Val(Cantidad)
                    
            ZLugar = A
            
            Erase Impre
                        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Laudo.Laudo"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                        
                With rstLaudo
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            If .EOF = True Then
                                Exit Do
                            End If
                           
                            dada = rstLaudo!Laudo
                           
                            WMarcaVencida = IIf(IsNull(rstLaudo!MarcaVencida), "", rstLaudo!MarcaVencida)
                            QSaldo = rstLaudo!Saldo
                            Call Redondeo(QSaldo)
                            If QSaldo <> 0 And Trim(WMarcaVencida) = "S" Then
                                If rstLaudo!Articulo = Articulo Then
                                    If Xlugar < 10 And XCanti > 0 Then
                                        Xlugar = Xlugar + 1
                                        Impre(Xlugar, 1) = rstLaudo!Laudo
                                        Impre(Xlugar, 2) = rstLaudo!Saldo
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
                    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Guia"
            ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
            ZSql = ZSql + " Order by Guia.Codigo"
            spMovguia = ZSql
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                           
                With rstMovguia
                           
                    .MoveFirst
                           
                    If .NoMatch = False Then
                        Do
                           
                            If .EOF = True Then
                                Exit Do
                            End If
                        
                            WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                            QSaldo = rstMovguia!Saldo
                            Call Redondeo(QSaldo)
                            If QSaldo <> 0 And Trim(WMarcaVencida) = "S" Then
                                If rstMovguia!Articulo = Articulo Then
                                    If Xlugar < 10 And XCanti > 0 Then
                                        Xlugar = Xlugar + 1
                                        Impre(Xlugar, 1) = rstMovguia!Lote
                                        Impre(Xlugar, 2) = rstMovguia!Saldo
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
                rstMovguia.Close
            End If
                    
            WArticulo1 = Left$(Articulo, 2)
            WArticulo2 = Mid$(Articulo, 4, 3) + "-" + Right$(Articulo, 3)
            WCantidad = Cantidad
                    
            For ZCiclo = 1 To 10
                ZZCanti(ZCiclo) = Str$(Impre(ZCiclo, 2))
                ZZLote(ZCiclo) = Str$(Impre(ZCiclo, 1))
            Next ZCiclo
                    
            Linea = Linea + 1
            WLinea = Str$(Linea)
                    
            For ZCiclo = 1 To 10
                    
                If Val(ZZCanti(ZCiclo)) <> 0 Then
                    
                    LineaII = LineaII + 1
                    WLIneaII = Str$(LineaII)
                        
                    WCantidadII = ZZCanti(ZCiclo)
                    WLoteII = ZZLote(ZCiclo)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM ImpreHojaII"
                    ZSql = ZSql + " Where ImpreHojaII.Lote = " + "'" + WLoteII + "'"

                    spImpreHojaII = ZSql
                    Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                    If rstImpreHojaII.RecordCount > 0 Then
                        rstImpreHojaII.Close
                            Else
                        ZSql = ""
                        ZSql = ZSql & "INSERT INTO ImpreHojaII ("
                        ZSql = ZSql & "Hoja ,"
                        ZSql = ZSql & "Renglon ,"
                        ZSql = ZSql & "Fecha ,"
                        ZSql = ZSql & "Codigo1 ,"
                        ZSql = ZSql & "Codigo2 ,"
                        ZSql = ZSql & "Maquina ,"
                        ZSql = ZSql & "Articulo1 ,"
                        ZSql = ZSql & "Articulo2 ,"
                        ZSql = ZSql & "Cantidad ,"
                        ZSql = ZSql & "Lote ,"
                        ZSql = ZSql & "Teorico ,"
                        ZSql = ZSql & "Terminado ,"
                        ZSql = ZSql & "Equipo )"
                        ZSql = ZSql & "Values ("
                        ZSql = ZSql & "'" + WHoja + "',"
                        ZSql = ZSql & "'" + WLIneaII + "',"
                        ZSql = ZSql & "'" + WFecha + "',"
                        ZSql = ZSql & "'" + WCodigo1 + "',"
                        ZSql = ZSql & "'" + WCodigo2 + "',"
                        ZSql = ZSql & "'" + WMaquina + "',"
                        ZSql = ZSql & "'" + WArticulo1 + "',"
                        ZSql = ZSql & "'" + WArticulo2 + "',"
                        ZSql = ZSql & "'" + WCantidadII + "',"
                        ZSql = ZSql & "'" + WLoteII + "',"
                        ZSql = ZSql & "'" + WTeorico + "',"
                        ZSql = ZSql & "'" + Producto.Text + "',"
                        ZSql = ZSql & "'" + Equipo.Text + "')"
        
                        spImpreHojaII = ZSql
                        Set rstImpreHojaII = db.OpenRecordset(spImpreHojaII, dbOpenSnapshot, dbSQLPassThrough)
                    
                        ZImpre = "S"
                    End If
                        
                End If
                        
            Next ZCiclo
                    
        End If
            
    Next A
    
    If ZImpre = "S" Then
    
    Listado.WindowTitle = "Impresion de Hoja de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Listado.ReportFileName = "ImpreHojaDesvio.rpt"
    Listado.SQLQuery = "SELECT ImpreHojaII.Hoja, ImpreHojaII.Renglon, ImpreHojaII.Fecha, ImpreHojaII.Articulo1, ImpreHojaII.Articulo2, ImpreHojaII.Cantidad, ImpreHojaII.Lote, ImpreHojaII.Terminado, ImpreHojaII.Equipo " _
                    + "From " _
                    + DSQ + ".dbo.ImpreHojaII ImprehojaII " _
                    + "Where ImpreHojaII.Hoja >= 0 AND ImpreHojaII.Hoja <= 999999"
                    
    Listado.GroupSelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"
    Listado.SelectionFormula = "{ImpreHojaII.Hoja} in 0 to 999999"

    Listado.Connect = Connect()
    Listado.Action = 1
    
    End If

End Sub

Private Sub Limpia_Grilla()

    Grilla.Clear

    Rem ponga la Grilla en negritas
    Grilla.Font.Bold = True

    ' Establesco loa Valores de la Grilla
    
    Grilla.FixedCols = 1
    Grilla.Cols = 7
    Grilla.FixedRows = 1
    Grilla.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem Grilla.Text = "Articulo"
    
    Rem Longitud
    Rem Grilla.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem Grilla.ColAlignment(Ciclo) = flexAlignRightCenter
    
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
    
    Grilla.ColWidth(0) = 200
    Grilla.Row = 0
    For Ciclo = 1 To Grilla.Cols - 1
        Grilla.Col = Ciclo
        Select Case Ciclo
            Case 1
                Grilla.Text = "Tipo"
                Grilla.ColWidth(Ciclo) = 500
                Grilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Grilla.Text = "Prod.Terminado"
                Grilla.ColWidth(Ciclo) = 1600
                Grilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Grilla.Text = "Materia Prima"
                Grilla.ColWidth(Ciclo) = 1500
                Grilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Grilla.Text = "Descripcion"
                Grilla.ColWidth(Ciclo) = 3600
                Grilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                Grilla.Text = "Cantidad"
                Grilla.ColWidth(Ciclo) = 1200
                Grilla.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                Grilla.Text = ""
                Grilla.ColWidth(Ciclo) = 10
                Grilla.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA Grilla
    
    WAncho = 400
    For Ciclo = 0 To Grilla.Cols - 1
        WAncho = WAncho + Grilla.ColWidth(Ciclo)
    Next Ciclo
    Rem Grilla.Width = WAncho

    ' Size the columns.
    Font.Name = Grilla.Font.Name
    Font.Size = Grilla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Grilla.AllowUserResizing = flexResizeBoth
    
    Grilla.Col = 1
    Grilla.Row = 1
    
End Sub

Private Sub Alta_Vector()

    ZLote(WWLugar, 1) = WLote1.Text
    ZLote(WWLugar, 2) = WCanti1.Text
    ZLote(WWLugar, 3) = WLote2.Text
    ZLote(WWLugar, 4) = WCanti2.Text
    ZLote(WWLugar, 5) = WLote3.Text
    ZLote(WWLugar, 6) = WCanti3.Text
    
    WLote1.Text = ""
    WCanti1.Text = ""
    WLote2.Text = ""
    WCanti2.Text = ""
    WLote3.Text = ""
    WCanti3.Text = ""
    WControl1.Locked = False
    WControl2.Locked = False
    WControl3.Locked = False
    WControl1.Text = ""
    WControl2.Text = ""
    WControl3.Text = ""
    WControl1.Locked = True
    WControl2.Locked = True
    WControl3.Locked = True
    
    CargaLote.Visible = False

End Sub














                
Private Sub GrabaVencido()

    Dim ZZVerifica(1000, 3) As String
    Dim ZZVerificaII(1000, 5) As String
    Dim ZZVerificaIII(1000, 20) As String
    Dim ZZVerificaIV(1000, 3) As String
    Dim ZZVerificaEmpe(1000, 3) As String
    
    Dim ZZVerificaLugar As Integer
    Dim ZZVerificaLugarII As Integer
    Dim ZZVerificaLugarIII As Integer
    Dim ZZVerificaLugarIV As Integer
    
    Dim VerificaSaldo As Double
    Dim ZZZVto As String
    Dim ZZZFechaVto As String
    
    
    ZZVeriticaArticulo = Producto.Text
    ZZEmpresaHoja = Wempresa
    ZZEmpresaPartida = 0
            
    Erase ZZVerifica
    Erase ZZVerificaII
    ZZVerificaLugar = 0
    ZZVerificaLugarII = 0
    
    Erase ZZVerificaIII
    ZZVerificaLugarIII = 0
        
    For A = 1 To 40
        ZZVerificaLugar = ZZVerificaLugar + 1
        ZZVerifica(ZZVerificaLugar, 1) = UCase(Grilla.TextMatrix(A, 1))
        ZZVerifica(ZZVerificaLugar, 2) = UCase(Grilla.TextMatrix(A, 2))
        ZZVerifica(ZZVerificaLugar, 3) = UCase(Grilla.TextMatrix(A, 3))
    Next A
            
    XEmpresa = Wempresa
    
    ZZVerificaEmpe(1, 1) = "0001"
    ZZVerificaEmpe(1, 2) = "Empresa01"
    ZZVerificaEmpe(2, 1) = "0003"
    ZZVerificaEmpe(2, 2) = "Empresa03"
    ZZVerificaEmpe(3, 1) = "0005"
    ZZVerificaEmpe(3, 2) = "Empresa05"
    ZZVerificaEmpe(4, 1) = "0006"
    ZZVerificaEmpe(4, 2) = "Empresa06"
    ZZVerificaEmpe(5, 1) = "0007"
    ZZVerificaEmpe(5, 2) = "Empresa07"
    ZZVerificaEmpe(6, 1) = "0010"
    ZZVerificaEmpe(6, 2) = "Empresa10"
    ZZVerificaEmpe(7, 1) = "0011"
    ZZVerificaEmpe(7, 2) = "Empresa11"
    
    ZZVerificaHasta = 7
            
    For VerificaCiclo = 1 To ZZVerificaHasta
    
        Wempresa = ZZVerificaEmpe(VerificaCiclo, 1)
        txtOdbc = ZZVerificaEmpe(VerificaCiclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
        For A = 1 To 40
            
            Tipo = UCase(ZZVerifica(A, 1))
            Terminado = UCase(ZZVerifica(A, 2))
            Articulo = UCase(ZZVerifica(A, 3))
                 
            Select Case Tipo
                
                Case "M"
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + Articulo + "'"
                    ZSql = ZSql + " and Laudo.Saldo <> 0"
                    ZSql = ZSql + " Order by Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                                
                        With rstLaudo
                            .MoveFirst
                            If .NoMatch = False Then
                                Do
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                   
                                    VerificaSaldo = rstLaudo!Saldo
                                    Call Redondeo(VerificaSaldo)
                                    If VerificaSaldo <> 0 Then
                                        
                                        ZZEntra = "S"
                                        For ZZMira = 1 To ZZVerificaLugarIII
                                            If ZZVerificaIII(ZZMira, 3) = Articulo And ZZVerificaIII(ZZMira, 4) = ZZVerificaLaudo Then
                                                ZZEntra = "N"
                                                Exit For
                                            End If
                                        Next ZZMira
                                        
                                        If ZZEntra = "S" Then
                                        
                                            ZZVerificaLaudo = rstLaudo!Laudo
                                                        
                                            If ZZVerificaLaudo <> 0 Then
                                                ZZVerificaLugarIII = ZZVerificaLugarIII + 1
                                                ZZVerificaIII(ZZVerificaLugarIII, 1) = "M"
                                                ZZVerificaIII(ZZVerificaLugarIII, 2) = "L"
                                                ZZVerificaIII(ZZVerificaLugarIII, 3) = Articulo
                                                ZZVerificaIII(ZZVerificaLugarIII, 4) = ZZVerificaLaudo
                                                ZZVerificaIII(ZZVerificaLugarIII, 5) = VerificaCiclo
                                                ZZVerificaIII(ZZVerificaLugarIII, 6) = rstLaudo!Fecha
                                                ZZVerificaIII(ZZVerificaLugarIII, 7) = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                                                ZZVerificaIII(ZZVerificaLugarIII, 18) = Str$(A)
                                                Select Case VerificaCiclo
                                                    Case 1
                                                        ZZVerificaIII(ZZVerificaLugarIII, 11) = Str$(VerificaSaldo)
                                                    Case 2
                                                        ZZVerificaIII(ZZVerificaLugarIII, 12) = Str$(VerificaSaldo)
                                                    Case 3
                                                        ZZVerificaIII(ZZVerificaLugarIII, 13) = Str$(VerificaSaldo)
                                                    Case 4
                                                        ZZVerificaIII(ZZVerificaLugarIII, 14) = Str$(VerificaSaldo)
                                                    Case 5
                                                        ZZVerificaIII(ZZVerificaLugarIII, 15) = Str$(VerificaSaldo)
                                                    Case 6
                                                        ZZVerificaIII(ZZVerificaLugarIII, 16) = Str$(VerificaSaldo)
                                                    Case 7
                                                        ZZVerificaIII(ZZVerificaLugarIII, 17) = Str$(VerificaSaldo)
                                                    Case Else
                                                End Select
                                            End If
                                            
                                                Else
                                                
                                                
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
                            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Guia"
                    ZSql = ZSql + " Where Guia.Articulo = " + "'" + Articulo + "'"
                    ZSql = ZSql + " and Guia.Saldo <> 0"
                    ZSql = ZSql + " Order by Guia.Codigo"
                    spMovguia = ZSql
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                                   
                        With rstMovguia
                                   
                            .MoveFirst
                                   
                            If .NoMatch = False Then
                                Do
                                   
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                                
                                    VerificaSaldo = rstMovguia!Saldo
                                    Call Redondeo(VerificaSaldo)
                                    If VerificaSaldo <> 0 Then
                                        
                                        ZZVerificaLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                                                    
                                        If ZZVerificaLaudo <> 0 Then
                                            ZZVerificaLugarIII = ZZVerificaLugarIII + 1
                                            ZZVerificaIII(ZZVerificaLugarIII, 1) = "M"
                                            ZZVerificaIII(ZZVerificaLugarIII, 2) = "G"
                                            ZZVerificaIII(ZZVerificaLugarIII, 3) = Articulo
                                            ZZVerificaIII(ZZVerificaLugarIII, 4) = ZZVerificaLaudo
                                            ZZVerificaIII(ZZVerificaLugarIII, 5) = ""
                                            ZZVerificaIII(ZZVerificaLugarIII, 6) = ""
                                            ZZVerificaIII(ZZVerificaLugarIII, 7) = ""
                                            ZZVerificaIII(ZZVerificaLugarIII, 18) = Str$(A)
                                    
                                            Select Case VerificaCiclo
                                                Case 1
                                                    ZZVerificaIII(ZZVerificaLugarIII, 11) = Str$(VerificaSaldo)
                                                Case 2
                                                    ZZVerificaIII(ZZVerificaLugarIII, 12) = Str$(VerificaSaldo)
                                                Case 3
                                                    ZZVerificaIII(ZZVerificaLugarIII, 13) = Str$(VerificaSaldo)
                                                Case 4
                                                    ZZVerificaIII(ZZVerificaLugarIII, 14) = Str$(VerificaSaldo)
                                                Case 5
                                                    ZZVerificaIII(ZZVerificaLugarIII, 15) = Str$(VerificaSaldo)
                                                Case 6
                                                    ZZVerificaIII(ZZVerificaLugarIII, 16) = Str$(VerificaSaldo)
                                                Case 7
                                                    ZZVerificaIII(ZZVerificaLugarIII, 17) = Str$(VerificaSaldo)
                                                Case Else
                                            End Select
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
                            
                                        
                Case "T"
                    XParam = "'" + Terminado + "','" _
                                 + Terminado + "'"
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
                                    
                                    VerificaSaldo = rstHoja!Saldo
                                    Call Redondeo(VerificaSaldo)
                                    
                                    If VerificaSaldo <> 0 And rstHoja!Renglon = 1 Then
                                        If rstHoja!Producto = Terminado Then
                                            
                                            ZZVerificaHoja = rstHoja!Hoja
                                                        
                                            If ZZVerificaHoja <> 0 Then
                                                ZZVerificaLugarIII = ZZVerificaLugarIII + 1
                                                ZZVerificaIII(ZZVerificaLugarIII, 1) = "T"
                                                ZZVerificaIII(ZZVerificaLugarIII, 2) = "H"
                                                ZZVerificaIII(ZZVerificaLugarIII, 3) = Terminado
                                                ZZVerificaIII(ZZVerificaLugarIII, 4) = ZZVerificaHoja
                                                ZZVerificaIII(ZZVerificaLugarIII, 5) = VerificaCiclo
                                                ZZVerificaIII(ZZVerificaLugarIII, 6) = rstHoja!Fecha
                                                ZZVerificaIII(ZZVerificaLugarIII, 7) = IIf(IsNull(rstHoja!FechaRevalida), "", rstHoja!FechaRevalida)
                                                ZZVerificaIII(ZZVerificaLugarIII, 8) = Str$(VerificaSaldo)
                                                ZZVerificaIII(ZZVerificaLugarIII, 9) = IIf(IsNull(rstHoja!Revalida), "", rstHoja!Revalida)
                                                ZZVerificaIII(ZZVerificaLugarIII, 10) = IIf(IsNull(rstHoja!MesesRevalida), "", rstHoja!MesesRevalida)
                                                ZZVerificaIII(ZZVerificaLugarIII, 18) = Str$(A)
                                                
                                                Select Case VerificaCiclo
                                                    Case 1
                                                        ZZVerificaIII(ZZVerificaLugarIII, 11) = Str$(VerificaSaldo)
                                                    Case 2
                                                        ZZVerificaIII(ZZVerificaLugarIII, 12) = Str$(VerificaSaldo)
                                                    Case 3
                                                        ZZVerificaIII(ZZVerificaLugarIII, 13) = Str$(VerificaSaldo)
                                                    Case 4
                                                        ZZVerificaIII(ZZVerificaLugarIII, 14) = Str$(VerificaSaldo)
                                                    Case 5
                                                        ZZVerificaIII(ZZVerificaLugarIII, 15) = Str$(VerificaSaldo)
                                                    Case 6
                                                        ZZVerificaIII(ZZVerificaLugarIII, 16) = Str$(VerificaSaldo)
                                                    Case 7
                                                        ZZVerificaIII(ZZVerificaLugarIII, 17) = Str$(VerificaSaldo)
                                                    Case Else
                                                End Select
                                                
                                                
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
                        rstHoja.Close
                    End If
                            
                    XParam = "'" + Terminado + "','" _
                                 + Terminado + "'"
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
                                
                                    VerificaSaldo = rstMovguia!Saldo
                                    Call Redondeo(VerificaSaldo)
                                    
                                    If VerificaSaldo <> 0 Then
                                        If rstMovguia!Terminado = Terminado Then
                                            
                                            ZZVerificaHoja = rstMovguia!Lote
                                                        
                                            If ZZVerificaHoja <> 0 Then
                                                ZZVerificaLugarIII = ZZVerificaLugarIII + 1
                                                ZZVerificaIII(ZZVerificaLugarIII, 1) = "T"
                                                ZZVerificaIII(ZZVerificaLugarIII, 2) = "G"
                                                ZZVerificaIII(ZZVerificaLugarIII, 3) = Terminado
                                                ZZVerificaIII(ZZVerificaLugarIII, 4) = ZZVerificaHoja
                                                ZZVerificaIII(ZZVerificaLugarIII, 5) = ""
                                                ZZVerificaIII(ZZVerificaLugarIII, 6) = ""
                                                ZZVerificaIII(ZZVerificaLugarIII, 7) = ""
                                                ZZVerificaIII(ZZVerificaLugarIII, 8) = Str$(VerificaSaldo)
                                                ZZVerificaIII(ZZVerificaLugarIII, 18) = Str$(A)
                                                Select Case VerificaCiclo
                                                    Case 1
                                                        ZZVerificaIII(ZZVerificaLugarIII, 11) = Str$(VerificaSaldo)
                                                    Case 2
                                                        ZZVerificaIII(ZZVerificaLugarIII, 12) = Str$(VerificaSaldo)
                                                    Case 3
                                                        ZZVerificaIII(ZZVerificaLugarIII, 13) = Str$(VerificaSaldo)
                                                    Case 4
                                                        ZZVerificaIII(ZZVerificaLugarIII, 14) = Str$(VerificaSaldo)
                                                    Case 5
                                                        ZZVerificaIII(ZZVerificaLugarIII, 15) = Str$(VerificaSaldo)
                                                    Case 6
                                                        ZZVerificaIII(ZZVerificaLugarIII, 16) = Str$(VerificaSaldo)
                                                    Case 7
                                                        ZZVerificaIII(ZZVerificaLugarIII, 17) = Str$(VerificaSaldo)
                                                    Case Else
                                                End Select
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
                        rstMovguia.Close
                    End If
            End Select
        
        Next A
    
    Next VerificaCiclo
    
    Call Conecta_Empresa
    
    For ZZZCiclaNuevo = 1 To ZZVerificaLugarIII
    
        ZZZTipo = ZZVerificaIII(ZZZCiclaNuevo, 1)
        ZZZOrigen = ZZVerificaIII(ZZZCiclaNuevo, 2)
        ZZZArticulo = ZZVerificaIII(ZZZCiclaNuevo, 3)
        ZZZLote = ZZVerificaIII(ZZZCiclaNuevo, 4)
        ZZZEmpresa = ZZVerificaIII(ZZZCiclaNuevo, 5)
        ZZZSaldoI = ZZVerificaIII(ZZZCiclaNuevo, 11)
        ZZZSaldoII = ZZVerificaIII(ZZZCiclaNuevo, 12)
        ZZZSaldoIII = ZZVerificaIII(ZZZCiclaNuevo, 13)
        ZZZSaldoIV = ZZVerificaIII(ZZZCiclaNuevo, 14)
        ZZZSaldoV = ZZVerificaIII(ZZZCiclaNuevo, 15)
        ZZZSaldoVI = ZZVerificaIII(ZZZCiclaNuevo, 16)
        ZZZSaldoVII = ZZVerificaIII(ZZZCiclaNuevo, 17)
        ZZZSaldo = Str$(Val(ZZZSaldoI) + Val(ZZZSaldoII) + Val(ZZZSaldoIII) + Val(ZZZSaldoIV) + Val(ZZZSaldoV) + Val(ZZZSaldoVI) + Val(ZZZSaldoVII))
        ZZZLugarGRilla = ZZVerificaIII(ZZZCiclaNuevo, 18)
        
        ZZZImpre = "N"
        If Val(ZZZSaldoI) <> 0 Then
            ZZZImpreI = "N"
                Else
            ZZZImpreI = ""
        End If
        If Val(ZZZSaldoII) <> 0 Then
            ZZZImpreII = "N"
                Else
            ZZZImpreII = ""
        End If
        If Val(ZZZSaldoIII) <> 0 Then
            ZZZImpreIII = "N"
                Else
            ZZZImpreIII = ""
        End If
        If Val(ZZZSaldoIV) <> 0 Then
            ZZZImpreIV = "N"
                Else
            ZZZImpreIV = ""
        End If
        If Val(ZZZSaldoV) <> 0 Then
            ZZZImpreV = "N"
                Else
            ZZZImpreV = ""
        End If
        If Val(ZZZSaldoVI) <> 0 Then
            ZZZImpreVI = "N"
                Else
            ZZZImpreVI = ""
        End If
        If Val(ZZZSaldoVII) <> 0 Then
            ZZZImpreVII = "N"
                Else
            ZZZImpreVII = ""
        End If
        
        Select Case ZZZTipo
            Case "M"
                ZZEntra = "S"
                If ZZZOrigen = "G" Then
                    
                    XEmpresa = Wempresa
    
                    ZZVerificaEmpe(1, 1) = "0001"
                    ZZVerificaEmpe(1, 2) = "Empresa01"
                    ZZVerificaEmpe(2, 1) = "0003"
                    ZZVerificaEmpe(2, 2) = "Empresa03"
                    ZZVerificaEmpe(3, 1) = "0005"
                    ZZVerificaEmpe(3, 2) = "Empresa05"
                    ZZVerificaEmpe(4, 1) = "0006"
                    ZZVerificaEmpe(4, 2) = "Empresa06"
                    ZZVerificaEmpe(5, 1) = "0007"
                    ZZVerificaEmpe(5, 2) = "Empresa07"
                    ZZVerificaEmpe(6, 1) = "0010"
                    ZZVerificaEmpe(6, 2) = "Empresa10"
                    ZZVerificaEmpe(7, 1) = "0011"
                    ZZVerificaEmpe(7, 2) = "Empresa11"
                    
                    ZZVerificaHasta = 7
                    ZZEntra = "N"
                            
                    For VerificaCiclo = 1 To ZZVerificaHasta
                    
                        Wempresa = ZZVerificaEmpe(VerificaCiclo, 1)
                        txtOdbc = ZZVerificaEmpe(VerificaCiclo, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + ZZZArticulo + "'"
                        ZSql = ZSql + " AND Laudo.Laudo = " + "'" + ZZZLote + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            ZZEntra = "S"
                            ZZVerificaIII(ZZZCiclaNuevo, 5) = VerificaCiclo
                            ZZVerificaIII(ZZZCiclaNuevo, 6) = rstLaudo!Fecha
                            ZZVerificaIII(ZZZCiclaNuevo, 7) = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                            rstLaudo.Close
                            Exit For
                        End If
                        
                    Next VerificaCiclo
                    
                    Call Conecta_Empresa
                    
                End If
                        
                If ZZEntra = "S" Then
                        
                    ZZZEmpresa = ZZVerificaIII(ZZZCiclaNuevo, 5)
                    ZZZFecha = ZZVerificaIII(ZZZCiclaNuevo, 6)
                    ZZZFechaVto = ZZVerificaIII(ZZZCiclaNuevo, 7)
    
                    ZZZVto = ""
                    ZZZMarcaVencida = ""
                            
                    ZZZOrdFecha = Right$(ZZZFecha, 4) + Mid$(ZZZFecha, 4, 2) + Left$(ZZZFecha, 2)
                    
                    If ZZZFechaVto <> "" And ZZZFechaVto <> "  /  /    " And ZZZFechaVto <> "00/00/0000" Then
                        Call Valida_fecha(ZZZFechaVto, Auxi)
                        If Auxi = "S" Then
                            ZZZVto = ZZZFechaVto
                        End If
                    End If
                            
                    If ZZZVto = "" Then
                            
                        ZZZMeses = 0
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZZZArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            ZZZMeses = rstArticulo!Meses
                            rstArticulo.Close
                        End If
                            
                        WMes = Val(Mid$(ZZZFecha, 4, 2))
                        WAno = Val(Right$(ZZZFecha, 4))
                        For ZCiclo = 1 To ZZZMeses
                            WMes = WMes + 1
                            If WMes > 12 Then
                                WAno = WAno + 1
                                WMes = 1
                            End If
                        Next ZCiclo
                            
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        If Val(Left$(ZZZFecha, 2)) <= 30 Then
                            If Val(XMes) = 2 And Val(Left$(ZZZFecha, 2)) > 28 Then
                                ZZZVto = "28/" + XMes + "/" + XAno
                                    Else
                                ZZZVto = Left$(ZZZFecha, 3) + XMes + "/" + XAno
                            End If
                                Else
                            If Val(XMes) = 2 Then
                                ZZZVto = "28/" + XMes + "/" + XAno
                                    Else
                                ZZZVto = "30/" + XMes + "/" + XAno
                            End If
                        End If
                           
                    End If
                        
                    If ZZZVto <> "" Then
                        
                        Do
                            Call Valida_fecha(ZZZVto, Auxi)
                            If Auxi = "S" Then
                                Exit Do
                                    Else
                                XFec1 = ZZZVto
                                SumaDia = 1
                                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                                ZZZVto = XFec2
                            End If
                        Loop
                        
                        ZZComparaI = Fecha.Text
                        If Left$(ZZZVto, 2) > "28" Then
                            ZZComparaII = "28" + Mid$(ZZZVto, 3, 8)
                                Else
                            ZZComparaII = ZZZVto
                        End If
                        
                        ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
                        
                        If Val(ZDias) < 0 Then
                            ZZZMarcaVencida = "S"
                        End If
                        
                    End If
                    
                    If ZZZMarcaVencida = "S" Then
                        
                        XEmpresa = Wempresa
                        
                        Select Case Val(ZZVerificaEmpe(ZZZEmpresa, 1))
                            Case 1
                                Wempresa = ZZVerificaEmpe(1, 1)
                                txtOdbc = ZZVerificaEmpe(1, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = ZZVerificaEmpe(2, 1)
                                txtOdbc = ZZVerificaEmpe(2, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZZZArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            ZZZDescripcion = rstArticulo!Descripcion
                            rstArticulo.Close
                        End If
                        
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM VerificaVtoArti"
                        ZSql = ZSql + " Where VerificaVtoArti.Articulo = " + "'" + ZZZArticulo + "'"
                        ZSql = ZSql + " and VerificaVtoArti.Partida = " + "'" + ZZZLote + "'"
                        
                        spVerificaVtoArti = ZSql
                        Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                        If rstVerificaVtoArti.RecordCount > 0 Then
                            rstVerificaVtoArti.Close
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE VerificaVtoArti SET "
                            ZSql = ZSql + " EmpresaTipo = " + "'" + ZZEmpresaHoja + "',"
                            ZSql = ZSql + " EmpresaPartida = " + "'" + ZZZEmpresa + "',"
                            ZSql = ZSql + " Estado = " + "'" + "0" + "',"
                            ZSql = ZSql + " Tipo = " + "'" + "HOJA" + "',"
                            ZSql = ZSql + " Numero = " + "'" + Hoja.Text + "',"
                            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
                            ZSql = ZSql + " Stock = " + "'" + ZZZSaldo + "',"
                            ZSql = ZSql + " StockI = " + "'" + ZZZSaldoI + "',"
                            ZSql = ZSql + " StockII = " + "'" + ZZZSaldoII + "',"
                            ZSql = ZSql + " StockIII = " + "'" + ZZZSaldoIII + "',"
                            ZSql = ZSql + " StockIV = " + "'" + ZZZSaldoIV + "',"
                            ZSql = ZSql + " StockV = " + "'" + ZZZSaldoV + "',"
                            ZSql = ZSql + " StockVI = " + "'" + ZZZSaldoVI + "',"
                            ZSql = ZSql + " StockVII = " + "'" + ZZZSaldoVII + "',"
                            ZSql = ZSql + " Impre = " + "'" + ZZZImpre + "',"
                            ZSql = ZSql + " ImpreOtro = " + "'" + ZZZImpre + "',"
                            ZSql = ZSql + " ImpreI = " + "'" + ZZZImpreI + "',"
                            ZSql = ZSql + " ImpreII = " + "'" + ZZZImpreII + "',"
                            ZSql = ZSql + " ImpreIII = " + "'" + ZZZImpreIII + "',"
                            ZSql = ZSql + " ImpreIV = " + "'" + ZZZImpreIV + "',"
                            ZSql = ZSql + " ImpreV = " + "'" + ZZZImpreV + "',"
                            ZSql = ZSql + " ImpreVI = " + "'" + ZZZImpreVI + "',"
                            ZSql = ZSql + " ImpreVII = " + "'" + ZZZImpreVII + "'"
                            ZSql = ZSql + " Where VerificaVtoArti.Articulo = " + "'" + ZZZArticulo + "'"
                            ZSql = ZSql + " and VerificaVtoArti.Partida = " + "'" + ZZZLote + "'"
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                            
                            
                                Else
                                
                            ZSql = ""
                            ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
                            ZSql = ZSql + " FROM VerificaVtoArti"
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                            If rstVerificaVtoArti.RecordCount > 0 Then
                                rstVerificaVtoArti.MoveLast
                                ZUltimo = IIf(IsNull(rstVerificaVtoArti!CodigoMayor), "0", rstVerificaVtoArti!CodigoMayor)
                                ZZCodigo = ZUltimo + 1
                                rstVerificaVtoArti.Close
                            End If
                                
                                
                            ZSql = ""
                            ZSql = ZSql & "INSERT INTO VerificaVtoArti ("
                            ZSql = ZSql & "Codigo ,"
                            ZSql = ZSql & "Articulo ,"
                            ZSql = ZSql & "Terminado ,"
                            ZSql = ZSql & "Descripcion ,"
                            ZSql = ZSql & "TipoMov ,"
                            ZSql = ZSql & "Partida ,"
                            ZSql = ZSql & "EmpresaTipo ,"
                            ZSql = ZSql & "EmpresaPartida ,"
                            ZSql = ZSql & "Tipo ,"
                            ZSql = ZSql & "Numero ,"
                            ZSql = ZSql & "Fecha ,"
                            ZSql = ZSql & "Stock ,"
                            ZSql = ZSql & "StockI ,"
                            ZSql = ZSql & "StockII ,"
                            ZSql = ZSql & "StockIII ,"
                            ZSql = ZSql & "StockIV ,"
                            ZSql = ZSql & "StockV ,"
                            ZSql = ZSql & "StockVI ,"
                            ZSql = ZSql & "StockVII ,"
                            ZSql = ZSql & "Estado ,"
                            ZSql = ZSql & "Impre ,"
                            ZSql = ZSql & "ImpreOtro ,"
                            ZSql = ZSql & "ImpreI ,"
                            ZSql = ZSql & "ImpreII ,"
                            ZSql = ZSql & "ImpreIII ,"
                            ZSql = ZSql & "ImpreIV ,"
                            ZSql = ZSql & "ImpreV ,"
                            ZSql = ZSql & "ImpreVI ,"
                            ZSql = ZSql & "ImpreVII )"
                            ZSql = ZSql & "Values ("
                            ZSql = ZSql & "'" + Str$(ZZCodigo) + "',"
                            ZSql = ZSql & "'" + ZZZArticulo + "',"
                            ZSql = ZSql & "'" + "  -     -   " + "',"
                            ZSql = ZSql & "'" + ZZZDescripcion + "',"
                            ZSql = ZSql & "'" + "M" + "',"
                            ZSql = ZSql & "'" + ZZZLote + "',"
                            ZSql = ZSql & "'" + ZZEmpresaHoja + "',"
                            ZSql = ZSql & "'" + ZZZEmpresa + "',"
                            ZSql = ZSql & "'" + "Hoja" + "',"
                            ZSql = ZSql & "'" + Hoja.Text + "',"
                            ZSql = ZSql & "'" + Fecha.Text + "',"
                            ZSql = ZSql & "'" + ZZZSaldo + "',"
                            ZSql = ZSql & "'" + ZZZSaldoI + "',"
                            ZSql = ZSql & "'" + ZZZSaldoII + "',"
                            ZSql = ZSql & "'" + ZZZSaldoIII + "',"
                            ZSql = ZSql & "'" + ZZZSaldoIV + "',"
                            ZSql = ZSql & "'" + ZZZSaldoV + "',"
                            ZSql = ZSql & "'" + ZZZSaldoVI + "',"
                            ZSql = ZSql & "'" + ZZZSaldoVII + "',"
                            ZSql = ZSql & "'" + "0" + "',"
                            ZSql = ZSql & "'" + ZZZImpre + "',"
                            ZSql = ZSql & "'" + ZZZImpre + "',"
                            ZSql = ZSql & "'" + ZZZImpreI + "',"
                            ZSql = ZSql & "'" + ZZZImpreII + "',"
                            ZSql = ZSql & "'" + ZZZImpreIII + "',"
                            ZSql = ZSql & "'" + ZZZImpreIV + "',"
                            ZSql = ZSql & "'" + ZZZImpreV + "',"
                            ZSql = ZSql & "'" + ZZZImpreVI + "',"
                            ZSql = ZSql & "'" + ZZZImpreVII + "')"
                
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                            
                        m$ = ZZZArticulo + " Venc.: "
                        If Val(ZZZSaldoI) <> 0 Then
                            m$ = m$ + ZZZSaldoI + " Kg PtaI "
                        End If
                        If Val(ZZZSaldoII) <> 0 Then
                            m$ = m$ + ZZZSaldoII + " Kg PtaII "
                        End If
                        If Val(ZZZSaldoIII) <> 0 Then
                            m$ = m$ + ZZZSaldoIII + " Kg PtaIII "
                        End If
                        If Val(ZZZSaldoIV) <> 0 Then
                            m$ = m$ + ZZZSaldoIV + " Kg PtaIV "
                        End If
                        If Val(ZZZSaldoV) <> 0 Then
                            m$ = m$ + ZZZSaldoV + " Kg PtaV "
                        End If
                        If Val(ZZZSaldoVI) <> 0 Then
                            m$ = m$ + ZZZSaldoVI + " Kg PtaVI "
                        End If
                        If Val(ZZZSaldoVII) <> 0 Then
                            m$ = m$ + ZZZSaldoVII + " Kg PtaVII "
                        End If
                        G% = MsgBox(m$, 0, "Verificacion de Partidas Vencidas")
                        
                        Grilla.TextMatrix(Val(ZZZLugarGRilla), 6) = m$
                        
                        Call Conecta_Empresa
                        
                    End If
                
                End If
                
            Case "T"
                ZZEntra = "S"
                If ZZZOrigen = "G" Then
                    
                    XEmpresa = Wempresa
    
                    ZZVerificaEmpe(1, 1) = "0001"
                    ZZVerificaEmpe(1, 2) = "Empresa01"
                    ZZVerificaEmpe(2, 1) = "0003"
                    ZZVerificaEmpe(2, 2) = "Empresa03"
                    ZZVerificaEmpe(3, 1) = "0005"
                    ZZVerificaEmpe(3, 2) = "Empresa05"
                    ZZVerificaEmpe(4, 1) = "0006"
                    ZZVerificaEmpe(4, 2) = "Empresa06"
                    ZZVerificaEmpe(5, 1) = "0007"
                    ZZVerificaEmpe(5, 2) = "Empresa07"
                    ZZVerificaEmpe(6, 1) = "0010"
                    ZZVerificaEmpe(6, 2) = "Empresa10"
                    ZZVerificaEmpe(7, 1) = "0011"
                    ZZVerificaEmpe(7, 2) = "Empresa11"
                    
                    ZZVerificaHasta = 7
                    ZZEntra = "N"
                            
                    For VerificaCiclo = 1 To ZZVerificaHasta
                    
                        Wempresa = ZZVerificaEmpe(VerificaCiclo, 1)
                        txtOdbc = ZZVerificaEmpe(VerificaCiclo, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                
                        spHoja = "ListaHoja " + "'" + ZZZLote + "'"
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                        
                            ZZEntra = "S"
                            ZZVerificaIII(ZZZCiclaNuevo, 5) = VerificaCiclo
                            ZZVerificaIII(ZZZCiclaNuevo, 6) = rstHoja!Fecha
                            ZZVerificaIII(ZZZCiclaNuevo, 7) = IIf(IsNull(rstHoja!FechaRevalida), "", rstHoja!FechaRevalida)
                            ZZVerificaIII(ZZZCiclaNuevo, 9) = IIf(IsNull(rstHoja!Revalida), "", rstHoja!Revalida)
                            ZZVerificaIII(ZZZCiclaNuevo, 10) = IIf(IsNull(rstHoja!MesesRevalida), "", rstHoja!MesesRevalida)
                            
                            rstHoja.Close
                            Exit For
                        End If
                        
                    Next VerificaCiclo
                    
                    Call Conecta_Empresa
                    
                End If
                        
                If ZZEntra = "S" Then
                    
                    ZZZEmpresa = ZZVerificaIII(ZZZCiclaNuevo, 5)
                    ZZZFecha = ZZVerificaIII(ZZZCiclaNuevo, 6)
                    ZZZFechaRevalida = ZZVerificaIII(ZZZCiclaNuevo, 7)
                    ZZZRevalida = ZZVerificaIII(ZZZCiclaNuevo, 9)
                    ZZZMesesRevalida = ZZVerificaIII(ZZZCiclaNuevo, 10)
                    
                    spTerminado = "ConsultaTerminado " + "'" + ZZZArticulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                        rstTerminado.Close
                    End If
                    
                    
                    WMes = Val(Mid$(ZZZFecha, 4, 2))
                    WAno = Val(Right$(ZZZFecha, 4))
                
                    If Val(ZZZRevalida) <> 0 Then
                        WMes = Val(Mid$(ZZZFechaRevalida, 4, 2))
                        WAno = Val(Right$(ZZZFechaRevalida, 4))
                        WVida = Val(ZZZMesesRevalida)
                    End If
                
                    For Ciclo = 1 To WVida
                        WMes = WMes + 1
                        If WMes > 12 Then
                            WAno = WAno + 1
                            WMes = 1
                        End If
                    Next Ciclo
                    WElaboracion = ZZZFecha
                    Rem XFec1 = WElaboracion
                    Rem SumaDia = WVida + 1
                    Rem Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                    If WVida <> 0 Then
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        ZZZVto = "01/" + XMes + "/" + XAno
                            Else
                        ZZZVto = Fecha.Text
                    End If
                
                    ZZComparaI = Fecha.Text
                    ZZComparaII = ZZZVto
                    ZDias = DateDiff("d", ZZComparaI, ZZComparaII)
                    
                    ZZZMarcaVencida = ""
                    If Val(ZDias) < 0 Then
                        ZZZMarcaVencida = "S"
                    End If
                
                
                
                
                
                    If ZZZMarcaVencida = "S" Then
                    
                        XEmpresa = Wempresa
    
                        Select Case Val(ZZVerificaEmpe(ZZZEmpresa, 1))
                            Case 1
                                Wempresa = ZZVerificaEmpe(1, 1)
                                txtOdbc = ZZVerificaEmpe(1, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                Wempresa = ZZVerificaEmpe(2, 1)
                                txtOdbc = ZZVerificaEmpe(2, 2)
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                        
                        spTerminado = "ConsultaTerminado " + "'" + ZZZArticulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZZZDescripcion = rstTerminado!Descripcion
                            rstTerminado.Close
                        End If
                        
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM VerificaVtoArti"
                        ZSql = ZSql + " Where VerificaVtoArti.Terminado = " + "'" + ZZZArticulo + "'"
                        ZSql = ZSql + " and VerificaVtoArti.Partida = " + "'" + ZZZLote + "'"
                        
                        spVerificaVtoArti = ZSql
                        Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                        If rstVerificaVtoArti.RecordCount > 0 Then
                            rstVerificaVtoArti.Close
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE VerificaVtoArti SET "
                            ZSql = ZSql + " EmpresaTipo = " + "'" + ZZEmpresaHoja + "',"
                            ZSql = ZSql + " EmpresaPartida = " + "'" + ZZZEmpresa + "',"
                            ZSql = ZSql + " Estado = " + "'" + "0" + "',"
                            ZSql = ZSql + " Tipo = " + "'" + "HOJA" + "',"
                            ZSql = ZSql + " Numero = " + "'" + Hoja.Text + "',"
                            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
                            ZSql = ZSql + " Stock = " + "'" + ZZZSaldo + "',"
                            ZSql = ZSql + " StockI = " + "'" + ZZZSaldoI + "',"
                            ZSql = ZSql + " StockII = " + "'" + ZZZSaldoII + "',"
                            ZSql = ZSql + " StockIII = " + "'" + ZZZSaldoIII + "',"
                            ZSql = ZSql + " StockIV = " + "'" + ZZZSaldoIV + "',"
                            ZSql = ZSql + " StockV = " + "'" + ZZZSaldoV + "',"
                            ZSql = ZSql + " StockVI = " + "'" + ZZZSaldoVI + "',"
                            ZSql = ZSql + " StockVII = " + "'" + ZZZSaldoVII + "',"
                            ZSql = ZSql + " Impre = " + "'" + ZZZImpre + "',"
                            ZSql = ZSql + " ImpreOtro = " + "'" + ZZZImpre + "',"
                            ZSql = ZSql + " ImpreI = " + "'" + ZZZImpreI + "',"
                            ZSql = ZSql + " ImpreII = " + "'" + ZZZImpreII + "',"
                            ZSql = ZSql + " ImpreIII = " + "'" + ZZZImpreIII + "',"
                            ZSql = ZSql + " ImpreIV = " + "'" + ZZZImpreIV + "',"
                            ZSql = ZSql + " ImpreV = " + "'" + ZZZImpreV + "',"
                            ZSql = ZSql + " ImpreVI = " + "'" + ZZZImpreVI + "',"
                            ZSql = ZSql + " ImpreVII = " + "'" + ZZZImpreVII + "'"
                            ZSql = ZSql + " Where VerificaVtoArti.Terminado = " + "'" + ZZZArticulo + "'"
                            ZSql = ZSql + " and VerificaVtoArti.Partida = " + "'" + ZZZLote + "'"
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                            
                            
                                Else
                                
                            ZSql = ""
                            ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
                            ZSql = ZSql + " FROM VerificaVtoArti"
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                            If rstVerificaVtoArti.RecordCount > 0 Then
                                rstVerificaVtoArti.MoveLast
                                ZUltimo = IIf(IsNull(rstVerificaVtoArti!CodigoMayor), "0", rstVerificaVtoArti!CodigoMayor)
                                ZZCodigo = ZUltimo + 1
                                rstVerificaVtoArti.Close
                            End If
                                
                                
                            ZSql = ""
                            ZSql = ZSql & "INSERT INTO VerificaVtoArti ("
                            ZSql = ZSql & "Codigo ,"
                            ZSql = ZSql & "Articulo ,"
                            ZSql = ZSql & "Terminado ,"
                            ZSql = ZSql & "Descripcion ,"
                            ZSql = ZSql & "TipoMov ,"
                            ZSql = ZSql & "Partida ,"
                            ZSql = ZSql & "EmpresaTipo ,"
                            ZSql = ZSql & "EmpresaPartida ,"
                            ZSql = ZSql & "Tipo ,"
                            ZSql = ZSql & "Numero ,"
                            ZSql = ZSql & "Fecha ,"
                            ZSql = ZSql & "Stock ,"
                            ZSql = ZSql & "StockI ,"
                            ZSql = ZSql & "StockII ,"
                            ZSql = ZSql & "StockIII ,"
                            ZSql = ZSql & "StockIV ,"
                            ZSql = ZSql & "StockV ,"
                            ZSql = ZSql & "StockVI ,"
                            ZSql = ZSql & "StockVII ,"
                            ZSql = ZSql & "Estado ,"
                            ZSql = ZSql & "Impre ,"
                            ZSql = ZSql & "ImpreOtro ,"
                            ZSql = ZSql & "ImpreI ,"
                            ZSql = ZSql & "ImpreII ,"
                            ZSql = ZSql & "ImpreIII ,"
                            ZSql = ZSql & "ImpreIV ,"
                            ZSql = ZSql & "ImpreV ,"
                            ZSql = ZSql & "ImpreVI ,"
                            ZSql = ZSql & "ImpreVII )"
                            ZSql = ZSql & "Values ("
                            ZSql = ZSql & "'" + Str$(ZZCodigo) + "',"
                            ZSql = ZSql & "'" + "  -   -   " + "',"
                            ZSql = ZSql & "'" + ZZZArticulo + "',"
                            ZSql = ZSql & "'" + ZZZDescripcion + "',"
                            ZSql = ZSql & "'" + "T" + "',"
                            ZSql = ZSql & "'" + ZZZLote + "',"
                            ZSql = ZSql & "'" + ZZEmpresaHoja + "',"
                            ZSql = ZSql & "'" + ZZZEmpresa + "',"
                            ZSql = ZSql & "'" + "Hoja" + "',"
                            ZSql = ZSql & "'" + Hoja.Text + "',"
                            ZSql = ZSql & "'" + Fecha.Text + "',"
                            ZSql = ZSql & "'" + ZZZSaldo + "',"
                            ZSql = ZSql & "'" + ZZZSaldoI + "',"
                            ZSql = ZSql & "'" + ZZZSaldoII + "',"
                            ZSql = ZSql & "'" + ZZZSaldoIII + "',"
                            ZSql = ZSql & "'" + ZZZSaldoIV + "',"
                            ZSql = ZSql & "'" + ZZZSaldoV + "',"
                            ZSql = ZSql & "'" + ZZZSaldoVI + "',"
                            ZSql = ZSql & "'" + ZZZSaldoVII + "',"
                            ZSql = ZSql & "'" + "0" + "',"
                            ZSql = ZSql & "'" + ZZZImpre + "',"
                            ZSql = ZSql & "'" + ZZZImpre + "',"
                            ZSql = ZSql & "'" + ZZZImpreI + "',"
                            ZSql = ZSql & "'" + ZZZImpreII + "',"
                            ZSql = ZSql & "'" + ZZZImpreIII + "',"
                            ZSql = ZSql & "'" + ZZZImpreIV + "',"
                            ZSql = ZSql & "'" + ZZZImpreV + "',"
                            ZSql = ZSql & "'" + ZZZImpreVI + "',"
                            ZSql = ZSql & "'" + ZZZImpreVII + "')"
                
                            spVerificaVtoArti = ZSql
                            Set rstVerificaVtoArti = db.OpenRecordset(spVerificaVtoArti, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                            
                        m$ = ZZZArticulo + "Venc.: "
                        If Val(ZZZSaldoI) <> 0 Then
                            m$ = m$ + ZZZSaldoI + " Kg PtaI "
                        End If
                        If Val(ZZZSaldoII) <> 0 Then
                            m$ = m$ + ZZZSaldoII + " Kg PtaII "
                        End If
                        If Val(ZZZSaldoIII) <> 0 Then
                            m$ = m$ + ZZZSaldoIII + " Kg PtaIII "
                        End If
                        If Val(ZZZSaldoIV) <> 0 Then
                            m$ = m$ + ZZZSaldoIV + " Kg PtaIV "
                        End If
                        If Val(ZZZSaldoV) <> 0 Then
                            m$ = m$ + ZZZSaldoV + " Kg PtaV "
                        End If
                        If Val(ZZZSaldoVI) <> 0 Then
                            m$ = m$ + ZZZSaldoVI + " Kg PtaVI "
                        End If
                        If Val(ZZZSaldoVII) <> 0 Then
                            m$ = m$ + ZZZSaldoVII + " Kg PtaVII "
                        End If
                        G% = MsgBox(m$, 0, "Verificacion de Partidas Vencidas")
                        
                        Grilla.TextMatrix(Val(ZZZLugarGRilla), 6) = m$
                        
                        Call Conecta_Empresa
                        
                    End If
                
                
                
                End If
                
            
            Case Else
        End Select
        
    Next ZZZCiclaNuevo
    
End Sub
