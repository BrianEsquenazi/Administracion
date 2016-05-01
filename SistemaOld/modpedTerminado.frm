VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModPedTerminado 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Pedidos "
   ClientHeight    =   7830
   ClientLeft      =   15
   ClientTop       =   510
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.ComboBox Via 
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
      Left            =   4920
      TabIndex        =   50
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Ficha_PtI 
      Caption         =   "Ficha  Planta V"
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
      Left            =   7080
      TabIndex        =   47
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Ficha_PtII 
      Caption         =   "   Ficha      Planta II"
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
      Left            =   8520
      TabIndex        =   46
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
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
      Index           =   11
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4200
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
      Index           =   10
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   3840
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
      Index           =   9
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Clave1 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3240
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   40
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   360
         Width           =   1815
      End
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector2 
      Height          =   615
      Left            =   1080
      TabIndex        =   32
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      _Version        =   327680
      BackColor       =   12648384
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3480
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3840
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3840
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   29
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   28
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
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2760
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
      TabIndex        =   26
      Top             =   2160
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   25
      Top             =   2760
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
      Left            =   3480
      TabIndex        =   24
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
      Index           =   4
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3120
      Width           =   375
   End
   Begin VB.Frame CargaLote 
      Caption         =   "Ingreso de Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   8040
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton ConfirmaCargaLote 
         Caption         =   "Confirma Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   45
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CommandButton CancelaCargaLote 
         Caption         =   "Cancela Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   36
         Top             =   2520
         Width           =   2775
      End
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
         Left            =   1920
         TabIndex        =   22
         Top             =   2040
         Width           =   1095
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
         Left            =   1920
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox WLote5 
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
         Left            =   480
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox WLote4 
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
         Left            =   480
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
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
         Left            =   1920
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
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
         Left            =   1920
         TabIndex        =   17
         Top             =   960
         Width           =   1095
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
         Left            =   1920
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Wlote3 
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
         Left            =   480
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox WLote2 
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
         Left            =   480
         MaxLength       =   10
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox WLote1 
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
         Left            =   480
         MaxLength       =   10
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cant."
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
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
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1215
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   120
      Width           =   1095
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
      Left            =   8520
      TabIndex        =   7
      Top             =   120
      Width           =   1335
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
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
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4080
      TabIndex        =   30
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
      Height          =   5415
      Left            =   0
      TabIndex        =   31
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9551
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label28 
      Caption         =   "Via"
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
      TabIndex        =   51
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label DesProducto 
      BackColor       =   &H00FFFFC0&
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
      Left            =   1320
      TabIndex        =   49
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      TabIndex        =   48
      Top             =   840
      Width           =   1095
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
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   4095
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
      TabIndex        =   4
      Top             =   480
      Width           =   1095
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgModPedTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WAceptada As String
Private WDirentrega As String
Private WFecEntrega As String
Private WDespago As String
Private WObservaciones As String

Private Auxiliar(100, 14) As String
Private ClavePedido(100) As String
Private BajaLote(5, 2) As String
Private xLote(100, 22) As String

Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstEnvase As Recordset
Dim spEnvase As String

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
Dim XCantidad1 As String
Dim xCantidad2 As String
Dim XLote1 As String
Dim XCantiLote1 As String
Dim XLote2 As String
Dim XCantiLote2 As String
Dim XLote3 As String
Dim XCantiLote3 As String
Dim XLote4 As String
Dim XCantiLote4 As String
Dim XLote5 As String
Dim XCantiLote5 As String

Dim XEnv1 As String
Dim XCantiEnv1 As String
Dim XEnv2 As String
Dim XCantiEnv2 As String
Dim XEnv3 As String
Dim XCantiEnv3 As String
Dim XEnv4 As String
Dim XCantiEnv4 As String
Dim XEnv5 As String
Dim XCantiEnv5 As String

Dim ControlLote(5, 2) As String
Dim WSaldo As Double
Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer
Dim WProceso As Integer
Dim ZSaldo As Double

Dim WGraba As String
Dim WTermi As String
Dim WArticulo As String

Dim ZLoteII(100, 30) As String
Dim ZLote(100, 5) As String
Dim ZCanti(100, 5) As String

Dim WWLote As String
Dim WWTipo As String

Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim WEspecif(100) As String

Dim ZAprueba As String
Dim ZOpcion(10) As Integer
    

Dim XEnvase(40, 6) As String
Dim Datos(100, 10) As String
Dim AuxiliarII(100, 5) As String
Dim WWCliente As String
Dim WWFecha  As String
Dim WWFecEntrega As String
Dim WWVersion As String
Dim WWTipoped As String
Dim WWObservaciones As String
Dim ZZLugarDirEntrega As Integer
Dim ZZDirEntrega(10) As String
Dim WWEspecif(100) As String
Dim WWRazon As String
Dim WWPago As String
Dim WWDirentrega As String
Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As Double
Dim WWPrecio As Double
Dim WWObserva As String
Dim WWOrdenCpa As String
Dim WWDesPago As String

Dim ZZRequiereCertificado As String
Dim ZZRequiereMsds As String
Dim ZZRequiereMsdsCada As String
Dim ZZRequiereHoja As String
Dim ZZPermiteParcial As String
Dim ZZPartidasVarias As String

Dim ZZEmailCertificado As String
Dim ZZEmailMsds As String
Dim ZZEmailHoja As String
Dim ZZDiasI As String
Dim ZZDiasII As String
Dim ZZDiasIII As String
Dim ZZEnvasesI As String
Dim ZZEnvasesII As String
Dim ZZEnvasesIII As String
Dim ZZEtiquetaI As String
Dim ZZEtiquetaII As String
Dim ZZEspecif1 As String
Dim ZZEspecif2 As String
Dim ZZEspecif3 As String
Dim ZZEspecif4 As String
Dim ZZEspecif5 As String
Dim ZZCantidadPartidas As String
Dim ImpreEnvase(10) As String

Dim ZZZTerminado As String
Dim ZZZLote As String
Dim ZZZPasa As String
Dim ZZRestriccion As Integer
Dim ZZRestriccionI As Integer
Dim ZZRestriccionII As Integer
Dim ZZVerifica(100, 2) As String
Dim CargaEmpresa(12, 2) As String
Dim ZHasta As Integer

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    
    ProcesoActivate = 1
    PrgModPedTerminado.Hide
    Unload Me
    PrgModifTerminado.Show
End Sub

Private Sub ConfirmaCargaLote_Click()

    If Val(WEmpresa) = 1 Then
    
        ZZRestriccion = 0
        ZZZPasa = "S"
        
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZRestriccion = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
            rstCliente.Close
        End If
        
        If ZZRestriccion = 1 Then
            
            ZZZTerminado = XTerminado
            
            If Trim(WLote1.Text) <> "" Then
                ZZZLote = WLote1.Text
                Call Verifica_Restriccion
                If ZZZPasa = "N" Then
                    Exit Sub
                End If
            End If
            
            If Trim(WLote2.Text) <> "" Then
                ZZZLote = WLote1.Text
                Call Verifica_Restriccion
                If ZZZPasa <> "N" Then
                    Exit Sub
                End If
            End If
            
            If Trim(Wlote3.Text) <> "" Then
                ZZZLote = WLote1.Text
                Call Verifica_Restriccion
                If ZZZPasa <> "N" Then
                    Exit Sub
                End If
            End If
            
            If Trim(WLote4.Text) <> "" Then
                ZZZLote = WLote1.Text
                Call Verifica_Restriccion
                If ZZZPasa <> "N" Then
                    Exit Sub
                End If
            End If
            
            If Trim(WLote5.Text) <> "" Then
                ZZZLote = WLote1.Text
                Call Verifica_Restriccion
                If ZZZPasa <> "N" Then
                    Exit Sub
                End If
            End If
            
        End If
    
    End If



    WLugar = WVector1.Row
    xLote(WLugar, 1) = WLote1.Text
    xLote(WLugar, 2) = WCanti1.Text
    xLote(WLugar, 3) = WLote2.Text
    xLote(WLugar, 4) = WCanti2.Text
    xLote(WLugar, 5) = Wlote3.Text
    xLote(WLugar, 6) = WCanti3.Text
    xLote(WLugar, 7) = WLote4.Text
    xLote(WLugar, 8) = WCanti4.Text
    xLote(WLugar, 9) = WLote5.Text
    xLote(WLugar, 10) = WCanti5.Text
    CargaLote.Visible = False
    Graba.Enabled = True
    If WVector1.Row < 40 Then
        WVector1.Row = WVector1.Row + 1
        WRow = WVector1.Row
        XRow = WVector1.Row
        WVector1.Col = 4
    End If
    WVector1.Row = XRow
    WVector1.Col = 3
    
End Sub

Private Sub Ficha_PtI_Click()
    Call ficha_Pt
End Sub

Private Sub Ficha_PtII_Click()
    Call Ficha_PtOtro
End Sub

Private Sub Graba_Click()
    
    Call Verifica_Certificado
    If ZAprueba = "N" Then
        Exit Sub
    End If




    Erase Auxiliar
    Auxi = 0
        
    Suma = 0
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 60
        
        WVector1.Row = a
                    
        WVector1.Col = 1
        Articulo = WVector1.Text
                    
        WVector1.Col = 4
        Cantidad = WVector1.Text
                    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        If Val(Cantidad) <> 0 Then
        
            XPedido = Left$(ClavePedido(a), 6)
            XRenglon = Right$(ClavePedido(a), 2)
            
            XParam = "'" + XPedido + "','" _
                     + XRenglon + "'"
            WClavePedido = ClavePedido(a)
            spPedido = "ConsultaPedido2 " + XParam
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                
                XCantidad1 = Cantidad
                xCantidad2 = Cantidad
                    
                WLugar = a
                
                XLote1 = xLote(WLugar, 1)
                XLote2 = xLote(WLugar, 3)
                XLote3 = xLote(WLugar, 5)
                XLote4 = xLote(WLugar, 7)
                XLote5 = xLote(WLugar, 9)
                
                XCantidad1 = Str$(Val(xLote(WLugar, 2)) + Val(xLote(WLugar, 4)) + Val(xLote(WLugar, 6)) + Val(xLote(WLugar, 8)) + Val(xLote(WLugar, 10)))
                xCantidad2 = Str$(Val(xLote(WLugar, 2)) + Val(xLote(WLugar, 4)) + Val(xLote(WLugar, 6)) + Val(xLote(WLugar, 8)) + Val(xLote(WLugar, 10)))
                
                XCantiLote1 = xLote(WLugar, 2)
                XCantiLote2 = xLote(WLugar, 4)
                XCantiLote3 = xLote(WLugar, 6)
                XCantiLote4 = xLote(WLugar, 8)
                XCantiLote5 = xLote(WLugar, 10)
                
                XEnv1 = ""
                XCantiEnv1 = ""
                XEnv2 = ""
                XCantiEnv2 = ""
                XEnv3 = ""
                XCantiEnv3 = ""
                XEnv4 = ""
                XCantiEnv4 = ""
                XEnv5 = ""
                XCantiEnv5 = ""
                XEti1 = ""
                XTipo1 = ""
                XEti2 = ""
                XTipo2 = ""
                XEti3 = ""
                XTipo3 = ""
                XEti4 = ""
                XTipo4 = ""
                XEti5 = ""
                XTipo5 = ""
                
                XParam = "'" + WClavePedido + "','" _
                         + XCantidad1 + "','" + xCantidad2 + "','" _
                         + XLote1 + "','" + XCantiLote1 + "','" _
                         + XLote2 + "','" + XCantiLote2 + "','" _
                         + XLote3 + "','" + XCantiLote3 + "','" _
                         + XLote4 + "','" + XCantiLote4 + "','" _
                         + XLote5 + "','" + XCantiLote5 + "','" _
                         + XEnv1 + "','" + XCantiEnv1 + "','" _
                         + XEnv2 + "','" + XCantiEnv2 + "','" _
                         + XEnv3 + "','" + XCantiEnv3 + "','" _
                         + XEnv4 + "','" + XCantiEnv4 + "','" _
                         + XEnv5 + "','" + XCantiEnv5 + "','" _
                         + XEti1 + "','" + XTipo1 + "','" _
                         + XEti2 + "','" + XTipo2 + "','" _
                         + XEti3 + "','" + XTipo3 + "','" _
                         + XEti4 + "','" + XTipo4 + "','" _
                         + XEti5 + "','" + XTipo5 + "'"
                                           
                spPedido = "ModificaPedidoActualiza " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                    
        End If
            
    Next a

    WPedido = Pedido.Text
    WWTipoped = ""

    Call ProcesoPedido_Click
    If Val(WWTipoped) = 5 Then
        Rem Call ImpresionIII
        Call ImpresionSql
            Else
        Call ImpresionSql
    End If

    WMarca = "X"
    XParam = "'" + WPedido + "','" _
            + WMarca + "'"
                               
    spPedido = "ModificaPedidoImpresion " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    WMarca = "N"
    XParam = "'" + WPedido + "','" _
            + WMarca + "'"
                               
    spPedido = "ModificaPedidoImpresion1 " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    WMarca = "3"
    XParam = "'" + WPedido + "','" _
                 + WMarca + "'"
                               
    spPedido = "ModificaPedidoProceso1 " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Call cmdClose_Click
    
End Sub

Private Sub Verifica_Certificado()

    WRenglon = 0
    ZAprueba = "S"


    For a = 1 To 60
        
        Articulo = WVector1.TextMatrix(a, 1)
                    
        If Trim(Articulo) <> "" Then
        
            WLugar = a
                
            WLote1 = xLote(WLugar, 1)
            WLote2 = xLote(WLugar, 3)
            Wlote3 = xLote(WLugar, 5)
            WLote4 = xLote(WLugar, 7)
            WLote5 = xLote(WLugar, 9)
        
            Rem
            Rem certificado de analisis
            Rem
        
            For ZZCiclo = 1 To 5
                
                Select Case ZZCiclo
                    Case 1
                        WWLote = WLote1
                    Case 2
                        WWLote = WLote2
                    Case 3
                        WWLote = Wlote3
                    Case 4
                        WWLote = WLote4
                    Case Else
                        WWLote = WLote5
                End Select
                
                If Trim(WWLote) <> "" Then
            
                    ZZEntra = "N"
            
                    If Left$(UCase(Articulo), 2) = "PT" Then
                    
                        XCodigo = Val(Mid$(Articulo, 4, 5))
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
                                        If XCodigo >= 40000 And XCodigo <= 41000 Then
                                            XTipoPro = "TA"
                                                Else
                                            XTipoPro = "PT"
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                        If Left$(Articulo, 2) = "YQ" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YH" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YP" Then
                            XTipoPro = "PT"
                        End If
                        If Left$(Articulo, 2) = "YF" Then
                            XTipoPro = "FA"
                        End If
                
                        ZLinea = 0
                        spTerminado = "ConsultaTerminado " + "'" + Articulo + "'"
                        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                        If rstTerminado.RecordCount > 0 Then
                            ZLinea = rstTerminado!Linea
                            rstTerminado.Close
                        End If
                
                        Select Case ZLinea
                            Case 8
                                XTipoPro = "PG"
                            Case 10, 20
                                XTipoPro = "FA"
                            Case Else
                        End Select
                
                        If XTipoPro <> "FA" And XTipoPro <> "TA" Then
                        
                            XEmpresa = WEmpresa
                            
                            Select Case Val(WEmpresa)
                                Case 1, 3, 5, 6, 7, 10, 11
                                    WEmpresa = "0003"
                                    txtOdbc = "Empresa03"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0004"
                                    txtOdbc = "Empresa04"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                            
                            ZProducto = Articulo
                                
                            ZZEntra = "N"
                            Erase ZOpcion
                            
                            ZSql = ""
                            ZSql = ZSql & "Select *"
                            ZSql = ZSql & " FROM AltaCertificado"
                            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + ZCliente + "'"
                            spAltaCertificado = ZSql
                            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstAltaCertificado.RecordCount > 0 Then
                                ZOpcion(1) = rstAltaCertificado!Opcion1
                                ZOpcion(2) = rstAltaCertificado!Opcion2
                                ZOpcion(3) = rstAltaCertificado!Opcion3
                                ZOpcion(4) = rstAltaCertificado!Opcion4
                                ZOpcion(5) = rstAltaCertificado!Opcion5
                                ZOpcion(6) = rstAltaCertificado!Opcion6
                                ZOpcion(7) = rstAltaCertificado!Opcion7
                                ZOpcion(8) = rstAltaCertificado!Opcion8
                                ZOpcion(9) = rstAltaCertificado!Opcion9
                                ZOpcion(10) = rstAltaCertificado!Opcion10
                                rstAltaCertificado.Close
                                ZZEntra = "S"
                            End If
                            
                            If ZZEntra = "N" Then
                                ZSql = ""
                                ZSql = ZSql & "Select *"
                                ZSql = ZSql & " FROM AltaCertificado"
                                ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + ZProducto + "'"
                                ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + "S00102" + "'"
                                spAltaCertificado = ZSql
                                Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstAltaCertificado.RecordCount > 0 Then
                                    ZOpcion(1) = rstAltaCertificado!Opcion1
                                    ZOpcion(2) = rstAltaCertificado!Opcion2
                                    ZOpcion(3) = rstAltaCertificado!Opcion3
                                    ZOpcion(4) = rstAltaCertificado!Opcion4
                                    ZOpcion(5) = rstAltaCertificado!Opcion5
                                    ZOpcion(6) = rstAltaCertificado!Opcion6
                                    ZOpcion(7) = rstAltaCertificado!Opcion7
                                    ZOpcion(8) = rstAltaCertificado!Opcion8
                                    ZOpcion(9) = rstAltaCertificado!Opcion9
                                    ZOpcion(10) = rstAltaCertificado!Opcion10
                                    rstAltaCertificado.Close
                                    ZZEntra = "S"
                                End If
                            End If
                            
                            Call Conecta_Empresa
                                
                            If ZZEntra = "S" Then
                                If ZOpcion(1) = 0 And ZOpcion(2) = 0 And ZOpcion(3) = 0 And ZOpcion(4) = 0 And ZOpcion(5) = 0 And ZOpcion(6) = 0 And ZOpcion(7) = 0 And ZOpcion(8) = 0 And ZOpcion(9) = 0 And ZOpcion(10) = 0 Then
                                    ZZEntra = "N"
                                End If
                            End If
                            
                            If ZZEntra = "N" Then
                                m$ = "El Certificado de Analisis de " + Articulo + " no se ha encontrado"
                                Aaa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                ZAprueba = "N"
                            End If
                                            
                        End If
                    
                            Else
                            
                        If Left$(Articulo, 2) = "DY" Then
                            
                            ZZPartiOri = Trim(WWLote)
                            ZZRuta = "w:\" + ZZPartiOri + ".PDF"
                            ZZEstado = Dir(ZZRuta)
                            ZZEstado = Trim(ZZEstado)
                            If ZZEstado = "" Then
                                m$ = "El articulo " + Articulo + " no tiene el certifiado de analisis de la partida " + ZZPartiOri
                                ssa% = MsgBox(m$, 0, "Imrpesion de comprobantes varios")
                                ZAprueba = "N"
                            End If
                            
                        End If
                        
                    End If
                End If
                    
            Next ZZCiclo
        End If
        
    Next a

End Sub


Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        WEntra = "N"
        
        XParam = "'" + WLote1.Text + "','" _
                + XTerminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            If WEstado <> "N" Then
                WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Else
                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                If WEstadoII = "V" Then
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Else
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                End If
                WSaldo1 = 0
            End If
            rstHoja.Close
        End If

        If WEntra = "N" Then
            XParam = "'" + XTerminado + "','" _
                    + WLote1.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                If WEstado <> "N" Then
                    WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Else
                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                    If WEstadoII = "V" Then
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            Else
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                    End If
                    WSaldo1 = 0
                End If
                rstMovguia.Close
            End If
        End If
            
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WLote1.Text = "" Then
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
                WVector1.Row = WVector1.Row + 1
                WRow = WVector1.Row
                XRow = WVector1.Row
                WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 3
            Exit Sub
        End If
            
        If WEntra = "S" Then
            ZZZTerminado = XTerminado
            ZZZLote = WLote1.Text
            Call Verifica_Restriccion
            If ZZZPasa <> "N" Then
                WCanti1.SetFocus
                    Else
                WSaldo1 = 0
            End If
                Else
            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
            G% = MsgBox(m$, 0, "Emision de facturas")
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WLote2.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote1.SetFocus
        End If
        Rem WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
        Rem WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WEntra = "N"
        
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        XParam = "'" + WLote2.Text + "','" _
                + XTerminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            If WEstado <> "N" Then
                WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Else
                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                If WEstadoII = "V" Then
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Else
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                End If
                WSaldo2 = 0
            End If
            rstHoja.Close
        End If
    
        If WEntra = "N" Then
            XParam = "'" + XTerminado + "','" _
                    + WLote2.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                If WEstado <> "N" Then
                    WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Else
                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                    If WEstadoII = "V" Then
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            Else
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                    End If
                    WSaldo2 = 0
                End If
                rstMovguia.Close
            End If
        End If
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WLote2.Text = "" Then
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
               WVector1.Row = WVector1.Row + 1
               WRow = WVector1.Row
               XRow = WVector1.Row
               WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 3
            Exit Sub
        End If
            
        If WEntra = "S" Then
            ZZZTerminado = XTerminado
            ZZZLote = WLote2.Text
            Call Verifica_Restriccion
            If ZZZPasa <> "N" Then
                WCanti2.SetFocus
                    Else
                WSaldo2 = 0
            End If
                Else
            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
            G% = MsgBox(m$, 0, "Emision de Facturas")
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            Wlote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote2.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        WEntra = "N"
        
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        XParam = "'" + Wlote3.Text + "','" _
                + XTerminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            If WEstado <> "N" Then
                WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Else
                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                If WEstadoII = "V" Then
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Else
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                End If
                WSaldo3 = 0
            End If
            rstHoja.Close
        End If
    
        If WEntra = "N" Then
            XParam = "'" + XTerminado + "','" _
                    + Wlote3.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                If WEstado <> "N" Then
                    WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Else
                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                    If WEstadoII = "V" Then
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            Else
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                    End If
                    WSaldo3 = 0
                End If
                rstMovguia.Close
            End If
        End If
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If Wlote3.Text = "" Then
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
               WVector1.Row = WVector1.Row + 1
               WRow = WVector1.Row
               XRow = WVector1.Row
               WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 3
            Exit Sub
        End If
            
        If WEntra = "S" Then
            ZZZTerminado = XTerminado
            ZZZLote = Wlote3.Text
            Call Verifica_Restriccion
            If ZZZPasa <> "N" Then
                WCanti3.SetFocus
                    Else
                WSaldo3 = 0
            End If
                Else
            m$ = XTerminado + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente"
            G% = MsgBox(m$, 0, "Emision de Facturas")
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            WLote4.SetFocus
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
            G% = MsgBox(m$, 0, "Emision de facturas")
            Wlote3.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WEntra = "N"
        
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        XParam = "'" + WLote4.Text + "','" _
                + XTerminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            If WEstado <> "N" Then
                WSaldo4 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Else
                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                If WEstadoII = "V" Then
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Else
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                End If
                WSaldo4 = 0
            End If
            rstHoja.Close
        End If
    
        If WEntra = "N" Then
            XParam = "'" + XTerminado + "','" _
                    + WLote4.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                If WEstado <> "N" Then
                    WSaldo4 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Else
                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                    If WEstadoII = "V" Then
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            Else
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                    End If
                    WSaldo4 = 0
                End If
                rstMovguia.Close
            End If
        End If
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WLote4.Text = "" Then
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
               WVector1.Row = WVector1.Row + 1
               WRow = WVector1.Row
               XRow = WVector1.Row
               WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 3
            Exit Sub
        End If
            
        If WEntra = "S" Then
            ZZZTerminado = XTerminado
            ZZZLote = WLote4.Text
            Call Verifica_Restriccion
            If ZZZPasa <> "N" Then
                WCanti4.SetFocus
                    Else
                WSaldo4 = 0
            End If
                Else
            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente"
            G% = MsgBox(m$, 0, "Emision de Facturas")
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo4 >= Val(WCanti4.Text) Then
            WCanti4.Text = Pusing("###,###.##", WCanti4.Text)
            WLote5.SetFocus
                Else
            XSaldo4 = WSaldo4
            XSaldo4 = Pusing("###,###.##", XSaldo4)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo4
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote4.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WEntra = "N"
        
        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        XParam = "'" + WLote5.Text + "','" _
                + XTerminado + "'"
        spHoja = "ListaHojaProducto " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            WEntra = "S"
            WEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
            If WEstado <> "N" Then
                WSaldo5 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                    Else
                WEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                If WEstadoII = "V" Then
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Else
                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                End If
                WSaldo5 = 0
            End If
            rstHoja.Close
        End If
    
        If WEntra = "N" Then
            XParam = "'" + XTerminado + "','" _
                    + WLote5.Text + "'"
            spMovguia = "ListaMovguiaLote1 " + XParam
            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovguia.RecordCount > 0 Then
                WEntra = "S"
                WEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                If WEstado <> "N" Then
                    WSaldo5 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Else
                    WEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                    If WEstadoII = "V" Then
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            Else
                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                    End If
                    WSaldo5 = 0
                End If
                rstMovguia.Close
            End If
        End If
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        If WLote5.Text = "" Then
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
               WVector1.Row = WVector1.Row + 1
               WRow = WVector1.Row
               XRow = WVector1.Row
               WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 3
            Exit Sub
        End If
            
        If WEntra = "S" Then
            ZZZTerminado = XTerminado
            ZZZLote = WLote5.Text
            Call Verifica_Restriccion
            If ZZZPasa = "S" Then
                WCanti5.SetFocus
                    Else
                WSaldo5 = 0
            End If
                Else
            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente"
            G% = MsgBox(m$, 0, "Emision de Facturas")
        End If
        
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo5 >= Val(WCanti5.Text) Then
            WCanti5.Text = Pusing("###,###.##", WCanti5.Text)
            WLugar = WVector1.Row
            xLote(WLugar, 1) = WLote1.Text
            xLote(WLugar, 2) = WCanti1.Text
            xLote(WLugar, 3) = WLote2.Text
            xLote(WLugar, 4) = WCanti2.Text
            xLote(WLugar, 5) = Wlote3.Text
            xLote(WLugar, 6) = WCanti3.Text
            xLote(WLugar, 7) = WLote4.Text
            xLote(WLugar, 8) = WCanti4.Text
            xLote(WLugar, 9) = WLote5.Text
            xLote(WLugar, 10) = WCanti5.Text
            CargaLote.Visible = False
            Graba.Enabled = True
            If WVector1.Row < 40 Then
                WVector1.Row = WVector1.Row + 1
                WRow = WVector1.Row
                XRow = WVector1.Row
                WVector1.Col = 4
            End If
            WVector1.Row = XRow
            WVector1.Col = 4
            Exit Sub
                Else
            XSaldo5 = WSaldo5
            XSaldo5 = Pusing("###,###.##", XSaldo5)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo5
            G% = MsgBox(m$, 0, "Emision de facturas")
            WLote5.SetFocus
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
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

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = 0
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Via.Clear
    
    Via.AddItem ""
    Via.AddItem "Terrestre"
    Via.AddItem "Maritimo"
    Via.AddItem "Aereo"
    
    Via.ListIndex = 0
    
    
    Pedido.Text = WXPed
    Call Pedido_KeyPress(13)
    
    Rem Pedido.SetFocus
     
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    Erase ClavePedido
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Canti = !Cantidad
                    
                    If Canti > 0 Then
                
                        Renglon = Renglon + 1
                        WVector1.Row = Renglon
                
                        WVector1.Col = 1
                        WVector1.Text = !Terminado
                        Auxi1 = !Terminado
                
                        WVector1.Col = 3
                        WVector1.Text = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                
                        WVector1.Col = 4
                        WVector1.Text = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                    
                        WLugar = Renglon
                        
                        xLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "", rstPedido!lote1)
                        xLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "", rstPedido!CantiLote1)
                        xLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "", rstPedido!lote2)
                        xLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "", rstPedido!CantiLote2)
                        xLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "", rstPedido!lote3)
                        xLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "", rstPedido!CantiLote3)
                        xLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "", rstPedido!lote4)
                        xLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "", rstPedido!CantiLote4)
                        xLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "", rstPedido!lote5)
                        xLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "", rstPedido!CantiLote5)
                        
                        Auxiliar(Renglon, 1) = Auxi1
                        Auxiliar(Renglon, 2) = Canti
                        
                        ClavePedido(Renglon) = rstPedido!Clave
                        
                    End If
        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(Da, 1)
        Canti = Auxiliar(Da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DW" Or Left$(Auxi1, 2) = "DS" Or Left$(Auxi1, 2) = "DQ" Then
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
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                For Ciclo = 1 To 9 Step 2
                    If Val(xLote(Da, Ciclo)) = 0 Then
                        xLote(Da, Ciclo) = ""
                            Else
                        ZEntra = "N"
                        Sql1 = "Select *"
                        Sql2 = " FROM Laudo"
                        Sql3 = " Where Laudo.Laudo = " + "'" + xLote(Da, Ciclo) + "'"
                        Sql4 = " and Laudo.Articulo = " + "'" + WArti + "'"
                        Sql5 = " Order by Laudo.Laudo"
                        spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            xLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            ZEntra = "S"
                            rstLaudo.Close
                        End If
                        
                        If ZEntra = "N" Then
                            Sql1 = "Select *"
                            Sql2 = " FROM Guia"
                            Sql3 = " Where Guia.Lote = " + "'" + xLote(Da, Ciclo) + "'"
                            Sql4 = " and Guia.Articulo = " + "'" + WArti + "'"
                            Sql5 = " Order by Guia.Saldo desc"
                            spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                xLote(Da, Ciclo) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                                ZEntra = "S"
                                rstMovguia.Close
                            End If
                        End If
                            
                        Rem XParam = "'" + xLote(Da, Ciclo) + "'"
                        Rem spLaudo = "ListaLaudo " + XParam
                        Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstLaudo.RecordCount > 0 Then
                        Rem     xLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                        Rem     rstLaudo.Close
                        Rem End If
                        
                    End If
                Next Ciclo
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WVector1.Row = Renglon
                    WVector1.Col = 2
                    WVector1.Text = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                For Ciclo = 1 To 9 Step 2
                    If Val(xLote(Da, Ciclo)) = 0 Then
                        xLote(Da, Ciclo) = ""
                    End If
                Next Ciclo
                
        End Select
        
    Next Da
    
    WVector1.TopRow = 1
    WVector1.Row = 1
    WVector1.Col = 1

End Sub

Private Sub Verifica_Lote()

    WEstado = "N"
    Suma = 0
    
    If WLote1.Text <> "" Then
        Suma = Suma + Val(WCanti1.Text)
    End If
    If WLote2.Text <> "" Then
        Suma = Suma + Val(WCanti2.Text)
    End If
    If Wlote3.Text <> "" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    If WLote4.Text <> "" Then
        Suma = Suma + Val(WCanti4.Text)
    End If
    If WLote5.Text <> "" Then
        Suma = Suma + Val(WCanti5.Text)
    End If
    
    If Suma = XCantidad Then
        WEstado = "S"
            Else
        m$ = "Las cantidades asignadas no concuerdan con las cantidades a facturar"
        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
    End If
    
    If WEstado = "S" Then
    
        Erase ControlLote
        ControlLote(1, 1) = WLote1.Text
        ControlLote(1, 2) = WCanti1.Text
        ControlLote(2, 1) = WLote2.Text
        ControlLote(2, 2) = WCanti2.Text
        ControlLote(3, 1) = Wlote3.Text
        ControlLote(3, 2) = WCanti3.Text
        ControlLote(4, 1) = WLote4.Text
        ControlLote(4, 2) = WCanti4.Text
        ControlLote(5, 1) = WLote5.Text
        ControlLote(5, 2) = WCanti5.Text
    
        For Ciclo1 = 1 To 5
            If Val(ControlLote(Ciclo1, 1)) <> 0 Then
                For Ciclo2 = 1 To 5
                    If Ciclo1 <> Ciclo2 Then
                        If Val(ControlLote(Ciclo1, 1)) = Val(ControlLote(Ciclo2, 1)) <> 0 Then
                            m$ = "A asignado una misma partida 2 veces"
                            a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                            WEstado = "N"
                            Exit For
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
        ControlLote(1, 1) = WLote1.Text
        ControlLote(1, 2) = WCanti1.Text
        ControlLote(2, 1) = WLote2.Text
        ControlLote(2, 2) = WCanti2.Text
        ControlLote(3, 1) = Wlote3.Text
        ControlLote(3, 2) = WCanti3.Text
        ControlLote(4, 1) = WLote4.Text
        ControlLote(4, 2) = WCanti4.Text
        ControlLote(5, 1) = WLote5.Text
        ControlLote(5, 2) = WCanti5.Text
    
        For Ciclo1 = 1 To 5
    
            WLote = ControlLote(Ciclo1, 1)
            WCanti = Val(ControlLote(Ciclo1, 2))
            
            If WLote <> "" Or Val(WCanti) <> 0 Then
            
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DW" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            Call Redondeo(WSaldo)
                            WEntra = "S"
                            If WSaldo < WCanti Then
                                m$ = "La cantidad informada supera al saldo disponible"
                                a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                WEstado = "N"
                            End If
                            ZEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            ZEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                            If ZEstado = "N" Then
                                If ZEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WEstado = "N"
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstMovguia.Close
                            End With
                        End If
                    End If
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
            
                    If WControla = 0 Then
                        XParam = "'" + WLote + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            Call Redondeo(WSaldo)
                            WEntra = "S"
                            If WSaldo < WCanti Then
                                m$ = "La cantidad informada supera al saldo disponible"
                                a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                WEstado = "N"
                            End If
                            ZEstado = IIf(IsNull(rstHoja!Estado), "", rstHoja!Estado)
                            ZEstadoII = IIf(IsNull(rstHoja!EstadoII), "", rstHoja!EstadoII)
                            If ZEstado = "N" Then
                                If ZEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WEstado = "N"
                            End If
                            WFechaHoja = rstHoja!Fecha
                            rstHoja.Close
                        End If
                
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                        + WLote + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                Call Redondeo(WSaldo)
                                WEntra = "S"
                                If WSaldo < WCanti Then
                                    m$ = "La cantidad informada supera al saldo disponible"
                                    a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                                    WEstado = "N"
                                End If
                                ZEstado = IIf(IsNull(rstMovguia!Estado), "", rstMovguia!Estado)
                                ZEstadoII = IIf(IsNull(rstMovguia!EstadoII), "", rstMovguia!EstadoII)
                                If ZEstado = "N" Then
                                    If ZEstadoII = "V" Then
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                            Else
                                        m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                        G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    End If
                                    WEstado = "N"
                                End If
                                rstMovguia.Close
                            End If
                        End If
                
                            Else
                        WEntra = "S"
                    End If
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
            End Select
            
            End If
            
        Next Ciclo1

    End If
    
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
        Rem Call Suma_Datos
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
            Rem If WControl = "S" Then
            Rem     Call Control_wvector1
            Rem End If
            Rem Call StartEdit
    
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
            WVector1.Col = 4
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
        Case 4
            WVector1.Col = 4
            Rem WVector1.Text = Pusing("###,###.##", Str$(Val(WVector1.Text)))
            WVector1.Text = WVector1.TextMatrix(Val(WVector1.Row), 3)
            WVector1.Col = 1
            XTerminado = WVector1.Text
            WVector1.Col = 3
            xcantidadoriginal = Val(WVector1.Text)
            WVector1.Col = 4
            XCantidad = Val(WVector1.Text)
            WRow = WVector1.Row
            
            Pasa = ""
            
            If XCantidad < xcantidadoriginal Then
                T$ = "MODIFICACION DE PEDIDOS"
                m$ = "ARTICULO = " + XTerminado + Chr$(13) + "CANTIDAD ORIGINAL DEL PEDIDO = " + Str$(xcantidadoriginal) + Chr$(13) + "CANTIDAD A INGRESAR = " + Str$(XCantidad) + Chr$(13) + "DIFERENCIA = " + Str$(xcantidadoriginal - XCantidad) + Chr$(13) + "ATENCION ! ! ! ! !   " + Chr$(13) + "LA DIFERENCIA ENTRE EL PEDIDO ORIGINAL Y LA CANTIDAD ACTUAL NO QUEDARA PENDIENTE DE ENTREGA" + Chr$(13) + "CONFIRMA ESTE PROCEDIMIENTO"
                Respuesta% = MsgBox(m$, 32 + 4 + 256, T$)
                If Respuesta% = 6 Then
                    Pasa = "S"
                End If
                    Else
                Pasa = "S"
            End If
            
            If Pasa = "S" Then
                CargaLote.Visible = True
                Graba.Enabled = False
                WLote1.Text = ""
                WCanti1.Text = ""
                WLote2.Text = ""
                WCanti2.Text = ""
                Wlote3.Text = ""
                WCanti3.Text = ""
                WLote4.Text = ""
                WCanti4.Text = ""
                WLote5.Text = ""
                WCanti5.Text = ""
                            
                WLugar = WVector1.Row
                
                If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DW" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                                
                    If xLote(WLugar, 1) <> "" Then
                        WLote1.Text = xLote(WLugar, 1)
                        WCanti1.Text = xLote(WLugar, 2)
                    End If
                    If xLote(WLugar, 3) <> "" Then
                        WLote2.Text = xLote(WLugar, 3)
                        WCanti2.Text = xLote(WLugar, 4)
                    End If
                    If xLote(WLugar, 5) <> "" Then
                        Wlote3.Text = xLote(WLugar, 5)
                        WCanti3.Text = xLote(WLugar, 6)
                    End If
                    If xLote(WLugar, 7) <> "" Then
                        WLote4.Text = xLote(WLugar, 7)
                        WCanti4.Text = xLote(WLugar, 8)
                    End If
                    If xLote(WLugar, 9) <> "" Then
                        WLote5.Text = xLote(WLugar, 9)
                        WCanti5.Text = xLote(WLugar, 10)
                    End If
                    
                        Else
                    
                    If Val(xLote(WLugar, 1)) <> 0 Then
                        WLote1.Text = xLote(WLugar, 1)
                        WCanti1.Text = xLote(WLugar, 2)
                    End If
                    If Val(xLote(WLugar, 3)) <> 0 Then
                        WLote2.Text = xLote(WLugar, 3)
                        WCanti2.Text = xLote(WLugar, 4)
                    End If
                    If Val(xLote(WLugar, 5)) <> 0 Then
                        Wlote3.Text = xLote(WLugar, 5)
                        WCanti3.Text = xLote(WLugar, 6)
                    End If
                    If Val(xLote(WLugar, 7)) <> 0 Then
                        WLote4.Text = xLote(WLugar, 7)
                        WCanti4.Text = xLote(WLugar, 8)
                    End If
                    If Val(xLote(WLugar, 9)) <> 0 Then
                        WLote5.Text = xLote(WLugar, 9)
                        WCanti5.Text = xLote(WLugar, 10)
                    End If
                
                End If
                WLote1.SetFocus
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 99 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 3
        WAuxi2 = WVector1.Text
        WVector1.Col = 4
        WAuxi3 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Or WAuxi3 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For Da = 1 To WVector1.Cols - 1
            WVector1.Col = Da
            WVector1.Text = WBorra(Ciclo, Da)
        Next Da
    Next Ciclo
    
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
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cant.Pedida"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 6
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 6
                WParametros(2, Ciclo) = 0
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

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                    Else
                Cliente.Text = rstPedido!Cliente
                Fecha.Text = rstPedido!Fecha
                WFecEntrega = rstPedido!FecEntrega
                WObservaciones = rstPedido!Observaciones
                ZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                Via.ListIndex = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                rstPedido.Close
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    
                    WDirentrega = rstCliente!DirEntrega
                    Rem WDirentrega = rstCliente!DirEntrega
                    WDirentrega = ""
        
                    ZDirEntrega(1) = rstCliente!DirEntrega
                    ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
                    ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
                    ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
                    ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
                    WDirentrega = ZDirEntrega(ZLugarDirEntrega)
                    
                    WPago = Str$(rstCliente!Pago1)
                    rstCliente.Close
                    spPago = "ConsultaPago " + "'" + WPago + "'"
                    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPago.RecordCount > 0 Then
                        WDespago = rstPago!Nombre
                        rstPago.Close
                    End If
                End If
                Call Proceso_Click
            End If
        End If
    End If
End Sub

Private Sub ficha_Pt()


    WEmpresa = "0007"
    txtOdbc = "Empresa07"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Call Limpia_Vector2
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    DesProducto.Caption = WTerminado
    XRenglon = 0
    
    Rem XParam = "'" + WTerminado + "','" _
    rem             + WTerminado + "'"
    Rem spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstHoja.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Hoja.Producto, Hoja.Marca, Hoja.Saldo, Hoja.Renglon, Hoja.Real, Hoja.Fecha, Hoja.Hoja, Hoja.MarcaVencida "
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Producto >= " + "'" + WTerminado + "'"
    ZSql = ZSql + " and Hoja.Producto <= " + "'" + WTerminado + "'"
    ZSql = ZSql + " and Hoja.Renglon = 1"
    ZSql = ZSql + " and Hoja.Saldo <> 0"
    ZSql = ZSql + " Order by Hoja.Hoja"
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
                    WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                        
                        If WMarcaVencida = "S" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HC0FFFF
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HC0FFFF
                            
                        End If
                        
                        Rem BY NAN
                        If WMarcaVencida = "V" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HFF&
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HFF&
                            
                        End If
                        
                
                        WVector2.Col = 1
                        WVector2.Text = "Hoja"
                        
                        WVector2.Col = 2
                        WVector2.Text = ZHoja
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = ZHoja
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
                        WVector2.Col = 9
                        WVector2.Text = ""
                    
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
    
        
    ZSql = ""
    ZSql = ZSql + "Select Guia.Terminado, Guia.Lote, Guia.Marca, Guia.Saldo, Guia.Terminado, Guia.Cantidad, Guia.Fecha, Guia.Codigo, Guia.Movi, Guia.Destino, Guia.TipoMov, Guia.Lote, Guia.Partida, Guia.MarcaVencida, Guia.Tipo"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
    ZSql = ZSql + " Order by Guia.Lote"
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    zterminado = rstMovguia!Terminado
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
                        WWTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        WWTipo = "Guia In"
                    End If
                    WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                        
                        If WMarcaVencida = "S" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HC0FFFF
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HC0FFFF
                            
                        End If
                        
                        Rem BY NAN
                        If WMarcaVencida = "V" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HFF&
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HFF&
                            
                        End If
                
                
                        WVector2.Col = 1
                        WVector2.Text = WWTipo
                        
                        WVector2.Col = 2
                        WVector2.Text = ZCodigo
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = WWLote
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
                        WVector2.Col = 9
                        WVector2.Text = ""
                        
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
    
    
    Rem XParam = "'" + WTerminado + "','" _
    rem              + WTerminado + "'"
    Rem spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstEntdev.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Entdev.Terminado, Entdev.Marca, Entdev.Terminado, Entdev.Cantidad, Entdev.Fecha, Entdev.Codigo, Entdev.Lote, Entdev.Saldo "
    ZSql = ZSql + " FROM Entdev"
    ZSql = ZSql + " Where Entdev.Terminado = " + "'" + WTerminado + "'"
    ZSql = ZSql + " Order by Entdev.Codigo"
    spEntdev = ZSql
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
                
                zterminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                WWLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector2.Row = XRenglon
                
                    WVector2.Col = 1
                    WVector2.Text = "Dev"
                        
                    WVector2.Col = 2
                    WVector2.Text = ZCodigo
                                               
                    WVector2.Col = 3
                    WVector2.Text = ZFecha
                        
                    WVector2.Col = 4
                    WVector2.Text = ""
                        
                    WVector2.Col = 5
                    WVector2.Text = ZCantidad
                
                    WVector2.Col = 6
                    WVector2.Text = ZSaldo
                
                    WVector2.Col = 7
                    WVector2.Text = WWLote
                        
                    WVector2.Col = 8
                    WVector2.Text = ""
                    
                    WVector2.Col = 9
                    WVector2.Text = ""

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
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
End Sub


Private Sub Ficha_PtOtro()

    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Call Limpia_Vector2
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    DesProducto.Caption = WTerminado
    XRenglon = 0
    
    Rem XParam = "'" + WTerminado + "','" _
    rem             + WTerminado + "'"
    Rem spHoja = "ListaHojaProductoDesdeHasta" + XParam
    Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstHoja.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Hoja.Producto, Hoja.Marca, Hoja.Saldo, Hoja.Renglon, Hoja.Real, Hoja.Fecha, Hoja.Hoja, Hoja.MarcaVencida "
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.Producto >= " + "'" + WTerminado + "'"
    ZSql = ZSql + " and Hoja.Producto <= " + "'" + WTerminado + "'"
    ZSql = ZSql + " and Hoja.Renglon = 1"
    ZSql = ZSql + " and Hoja.Saldo <> 0"
    ZSql = ZSql + " Order by Hoja.Hoja"
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
                    WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                    
                    If ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                        
                        If WMarcaVencida = "S" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HC0FFFF
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HC0FFFF
                            
                        End If
                        
                        Rem BY NAN
                        If WMarcaVencida = "V" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HFF&
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HFF&
                            
                        End If
                        
                
                        WVector2.Col = 1
                        WVector2.Text = "Hoja"
                        
                        WVector2.Col = 2
                        WVector2.Text = ZHoja
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = ZHoja
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
                        WVector2.Col = 9
                        WVector2.Text = ""
                    
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
    
        
    ZSql = ""
    ZSql = ZSql + "Select Guia.Terminado, Guia.Lote, Guia.Marca, Guia.Saldo, Guia.Terminado, Guia.Cantidad, Guia.Fecha, Guia.Codigo, Guia.Movi, Guia.Destino, Guia.TipoMov, Guia.Lote, Guia.Partida, Guia.MarcaVencida, Guia.Tipo"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
    ZSql = ZSql + " Order by Guia.Lote"
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    zterminado = rstMovguia!Terminado
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
                        WWTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        WWTipo = "Guia In"
                    End If
                    WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                        
                        If WMarcaVencida = "S" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HC0FFFF
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HC0FFFF
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HC0FFFF
                            
                        End If
                        
                        Rem BY NAN
                        If WMarcaVencida = "V" Then
                             
                             WVector2.Col = 1
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 2
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 3
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 4
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 5
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 6
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 7
                             WVector2.CellBackColor = &HFF&
                            
                             WVector2.Col = 8
                             WVector2.CellBackColor = &HFF&
                             
                             WVector2.Col = 9
                             WVector2.CellBackColor = &HFF&
                            
                        End If
                
                
                        WVector2.Col = 1
                        WVector2.Text = WWTipo
                        
                        WVector2.Col = 2
                        WVector2.Text = ZCodigo
                                               
                        WVector2.Col = 3
                        WVector2.Text = ZFecha
                        
                        WVector2.Col = 4
                        WVector2.Text = ""
                        
                        WVector2.Col = 5
                        WVector2.Text = ZCantidad
                
                        WVector2.Col = 6
                        WVector2.Text = ZSaldo
                
                        WVector2.Col = 7
                        WVector2.Text = WWLote
                        
                        WVector2.Col = 8
                        WVector2.Text = ""
                        
                        WVector2.Col = 9
                        WVector2.Text = ""
                        
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
    
    
    Rem XParam = "'" + WTerminado + "','" _
    rem              + WTerminado + "'"
    Rem spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Rem Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstEntdev.RecordCount > 0 Then
    
    ZSql = ""
    ZSql = ZSql + "Select Entdev.Terminado, Entdev.Marca, Entdev.Terminado, Entdev.Cantidad, Entdev.Fecha, Entdev.Codigo, Entdev.Lote, Entdev.Saldo "
    ZSql = ZSql + " FROM Entdev"
    ZSql = ZSql + " Where Entdev.Terminado = " + "'" + WTerminado + "'"
    ZSql = ZSql + " Order by Entdev.Codigo"
    spEntdev = ZSql
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
                
                zterminado = rstEntdev!Terminado
                ZCantidad = rstEntdev!Cantidad
                ZFecha = rstEntdev!Fecha
                ZCodigo = rstEntdev!Codigo
                WWLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                ZSaldo = rstEntdev!Saldo
                Call Redondeo(ZSaldo)
                
                If ZSaldo <> 0 Then
                    
                    XRenglon = XRenglon + 1
                    WVector2.Row = XRenglon
                
                    WVector2.Col = 1
                    WVector2.Text = "Dev"
                        
                    WVector2.Col = 2
                    WVector2.Text = ZCodigo
                                               
                    WVector2.Col = 3
                    WVector2.Text = ZFecha
                        
                    WVector2.Col = 4
                    WVector2.Text = ""
                        
                    WVector2.Col = 5
                    WVector2.Text = ZCantidad
                
                    WVector2.Col = 6
                    WVector2.Text = ZSaldo
                
                    WVector2.Col = 7
                    WVector2.Text = WWLote
                        
                    WVector2.Col = 8
                    WVector2.Text = ""
                    
                    WVector2.Col = 9
                    WVector2.Text = ""

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
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
End Sub


Private Sub Limpia_Vector2()

    WVector2.Height = 4095
    WVector2.Left = 120
    WVector2.Top = 1350
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 10
    WVector2.FixedRows = 1
    WVector2.Rows = 5001
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Saldo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WVector2.Text = "Cant.Ped."
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Disponible"
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTitulo(Ciclo).Text = WVector2.Text
        WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTitulo(Ciclo).Width = WVector2.CellWidth
        WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Visible = True
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub Limpia_Vector2II()

    WVector2.Height = 4095
    WVector2.Left = 120
    WVector2.Top = 1350
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 10
    WVector2.FixedRows = 1
    WVector2.Rows = 5001
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Tipo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Numero"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WVector2.Text = "Orden"
                WVector2.ColWidth(Ciclo) = 10
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Saldo"
                WVector2.ColWidth(Ciclo) = 800
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WVector2.Text = "Cant.Ped."
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Disponible"
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTitulo(Ciclo).Text = WVector2.Text
        WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTitulo(Ciclo).Width = WVector2.CellWidth
        WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Visible = True
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub


Private Sub WVector2_Click()
    busquedalote = WVector2.TextMatrix(WVector2.Row, 7)
    WVector2.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    WTitulo(5).Visible = False
    WTitulo(6).Visible = False
    WTitulo(7).Visible = False
    WTitulo(8).Visible = False
    WTitulo(9).Visible = False
    WTitulo(10).Visible = False
    WTitulo(11).Visible = False
    Ficha_PtI.Visible = False
    Ficha_PtII.Visible = False
    Select Case WProceso
        Case 1
            WLote1.Text = busquedalote
            Call Wlote1_Keypress(13)
        Case 2
            WLote2.Text = busquedalote
            Call Wlote2_Keypress(13)
        Case 3
            Wlote3.Text = busquedalote
            Call Wlote3_Keypress(13)
        Case 4
            WLote4.Text = busquedalote
            Call Wlote4_Keypress(13)
        Case 5
            WLote5.Text = busquedalote
            Call Wlote5_Keypress(13)
        Case Else
    End Select
        
End Sub

Private Sub WLote1_DblClick()
    WProceso = 1
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Call ficha_Pt
End Sub

Private Sub WLote2_DblClick()
    WProceso = 2
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Call ficha_Pt
End Sub

Private Sub WLote3_DblClick()
    WProceso = 3
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Call ficha_Pt
End Sub

Private Sub WLote4_DblClick()
    WProceso = 4
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Call ficha_Pt
End Sub

Private Sub WLote5_DblClick()
    WProceso = 5
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Call ficha_Pt
End Sub

Private Sub CancelaCargaLote_Click()
    CargaLote.Visible = False
    Graba.Enabled = True
End Sub

Private Sub WTexto2_DblClick()
    WProceso = 0
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    Ficha_PtI.Visible = True
    Ficha_PtII.Visible = True
    Call ficha_Pt
End Sub

Private Sub ProcesoPedido_Click()

    Erase XEnvase
    Erase Datos
    Erase AuxiliarII
    
    Renglon = 0
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WWCliente = rstPedido!Cliente
                    WWFecha = rstPedido!Fecha
                    WWFecEntrega = rstPedido!FecEntrega
                    WWVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                    WWTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                    WWObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                    WWObservaciones = Left$(WObservaciones + Space$(100), 100)
                    ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
                    Via.ListIndex = IIf(IsNull(rstPedido!Via), "0", rstPedido!Via)
                    WWOrdenCpa = IIf(IsNull(rstPedido!OrdenCpa), "", rstPedido!OrdenCpa)
                    
                    Rem If rstPedido!Cantidad - rstPedido!Facturado > 0 Then
                    If rstPedido!Cantidad > 0 Then
            
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", Str$(rstPedido!Cantidad - rstPedido!Facturado))
                        Rem Datos(Renglon, 2) = Pusing("###,###.##", Str$(rstPedido!Cantidad))
                        Datos(Renglon, 4) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        Datos(Renglon, 5) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        Datos(Renglon, 6) = IIf(IsNull(rstPedido!proceso2), "", rstPedido!proceso2)
                        Datos(Renglon, 7) = IIf(IsNull(rstPedido!cantiproceso), "", rstPedido!cantiproceso)
                        Datos(Renglon, 8) = IIf(IsNull(rstPedido!observa), "", rstPedido!observa)
                        Datos(Renglon, 9) = IIf(IsNull(rstPedido!Especificaciones), "", rstPedido!Especificaciones)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        AuxiliarII(WRenglon, 1) = rstPedido!Cliente
                        AuxiliarII(WRenglon, 2) = rstPedido!Terminado
                        AuxiliarII(WRenglon, 3) = IIf(IsNull(rstPedido!Tipopro), "", rstPedido!Tipopro)
                        AuxiliarII(WRenglon, 4) = IIf(IsNull(rstPedido!Articulo), "", rstPedido!Articulo)
                        If Left$(rstPedido!Terminado, 2) = "ML" Then
                            AuxiliarII(WRenglon, 5) = IIf(IsNull(rstPedido!NombreComercial), "", rstPedido!NombreComercial)
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
    
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Cliente = AuxiliarII(Da, 1)
        Terminado = AuxiliarII(Da, 2)
        Tipopro = AuxiliarII(Da, 3)
        Articulo = AuxiliarII(Da, 4)
        ZZNombreComercial = AuxiliarII(Da, 5)
        
        Renglon = Renglon + 1
        
        If Left$(Terminado, 2) = "PT" Or Left$(Terminado, 2) = "YQ" Or Left$(Terminado, 2) = "YF" Then
            spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Datos(Renglon, 1) = rstPrecios!Descripcion
                Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
                rstPrecios.Close
            End If
                Else
            spPreciosMp = "ConsultaPreciosMp " + "'" + Cliente + Articulo + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                Datos(Renglon, 3) = Pusing("###,###.##", rstPreciosMp!Precio)
                rstPreciosMp.Close
            End If
            
            If ZZNombreComercial <> "" Then
                Datos(Renglon, 1) = ZZNombreComercial
                    Else
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Datos(Renglon, 1) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            End If
        End If
    Next Da

End Sub

Private Sub Impresion()

    Open "dada.txt" For Output As #1
    Rem Open "lpt1" For Output As #1

    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        WWRazon = rstCliente!Razon
        WWDirentrega = ""
        
        ZZDirEntrega(1) = rstCliente!DirEntrega
        ZZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WWDirentrega = ZZDirEntrega(ZZLugarDirEntrega)
        
        Erase WWEspecif
        WWEspecif(1) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
        WWEspecif(2) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
        WWEspecif(3) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
        WWEspecif(4) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
        WWEspecif(5) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
        WWEspecif(6) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
        WWEspecif(7) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
        WWEspecif(8) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
        WWEspecif(9) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
        WWEspecif(10) = IIf(IsNull(rstCliente!Especif10), "", rstCliente!Especif10)
        For CicloEspecif = 1 To 10
            WWEspecif(CicloEspecif) = RTrim(WWEspecif(CicloEspecif))
        Next CicloEspecif
        
        rstCliente.Close
    End If
    
    WVia = ""
    Select Case Via.ListIndex
        Case 1
            WVia = "Pedido Exportacion : " + "Terrestre"
        Case 2
            WVia = "Pedido Exportacion : " + "Maritimo"
        Case 3
            WVia = "Pedido Exportacion : " + "Aereo"
        Case Else
    End Select
    

    For XX = 1 To 1

        Print #1, Tab(1); String$(79, "-")
        
        If Val(WWTipoped) = 5 Then
            Print #1, Tab(1); "|                         MUESTRAS PARA CLIENTE";
            Print #1, Tab(80); "|"
                Else
            Print #1, Tab(1); "|                         SURFACTAN S.A.";
            Print #1, Tab(80); "|"
        End If
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Pedido.Text;
        Print #1, " / ";
        Print #1, WWVersion; "  "; WVia;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Cliente";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Cliente.Text;
        Print #1, Tab(40); Left$(WWRazon, 35);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WWFecha;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Ent.";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WWFecEntrega;
        Select Case Val(WWTipoped)
            Case 0
                Print #1, " (Normal)";
            Case 1
                Print #1, " (A fecha)";
            Case 2
                Print #1, " (Fecha Limite)";
            Case 3
                Print #1, " (Urgente)";
            Case 4
                Print #1, " (Retira Cliente)";
            Case 5
                Print #1, " (Muestra)";
            Case Else
        End Select
            
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Entrega";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WWDirentrega;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Left$(WWObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Right$(WWObservaciones, 50);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|";
        Print #1, Tab(2); "Producto";
        Print #1, Tab(16); "|";
        Print #1, Tab(17); "Descripcion";
        Print #1, Tab(40); "|";
        Print #1, Tab(41); "Pedido";
        Print #1, Tab(50); "|";
        Print #1, Tab(51); "Partida";
        Print #1, Tab(58); "|";
        Print #1, Tab(59); " Cant.";
        Print #1, Tab(67); "|";
        Print #1, Tab(68); "Envase";
        Print #1, Tab(80); "|"

        Print #1, Tab(1); String$(79, "-")
        
        XLinea = 0
        WCounter = 0
                    
        For WCounter = 1 To 40
        
            If Datos(WCounter, 0) <> "" Then
                    
                WWArticulo = Datos(WCounter, 0)
                WWDescripcion = Datos(WCounter, 1)
                WWCantidad = Val(Datos(WCounter, 2))
                WWPrecio = Val(Datos(WCounter, 3))
                WWObserva = Datos(WCounter, 8)
                WWEspecifica = Datos(WCounter, 9)
                    
                Print #1, Tab(1); "|";
                Print #1, Tab(2); WWArticulo;
                Print #1, Tab(16); "|";
                Print #1, Tab(17); Left$(WWDescripcion, 22);
                Print #1, Tab(40); "|";
                Print #1, Tab(41); Pusing("###,###", Str$(WWCantidad));
                Print #1, Tab(50); "|";
                If Val(xLote(WCounter, 1)) <> 0 Then
                    Print #1, Tab(51); Pusing("######", xLote(WCounter, 1));
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Pusing("###,###", xLote(WCounter, 2));
                End If
                Print #1, Tab(67); "|";
                
                spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, 1) + "'"
                Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnvase.RecordCount > 0 Then
                    WAbre = rstEnvase!Abreviatura
                    rstEnvase.Close
                        Else
                    WAbre = ""
                End If

                Print #1, Tab(68); Alinea("###", XEnvase(WCounter, 2)) + " " + Left$(WAbre, 8);
                Print #1, Tab(80); "|"
                
                XLinea = XLinea + 1
                
                If Val(xLote(WCounter, 3)) <> 0 Or Val(XEnvase(WCounter, 3)) <> 0 Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|";
                    Print #1, Tab(40); "|";
                    Print #1, Tab(50); "|";
                    Print #1, Tab(51); Pusing("######", xLote(WCounter, 3));
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Pusing("###,###", xLote(WCounter, 4));
                    Print #1, Tab(67); "|";
                    
                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, 3) + "'"
                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvase.RecordCount > 0 Then
                        WAbre = rstEnvase!Abreviatura
                        rstEnvase.Close
                            Else
                        WAbre = ""
                    End If
        
                    Print #1, Tab(68); Alinea("###", XEnvase(WCounter, 4)) + " " + Left$(WAbre, 8);
                    
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                If Val(xLote(WCounter, 5)) <> 0 Or Val(XEnvase(WCounter, 5)) <> 0 Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|";
                    Print #1, Tab(40); "|";
                    Print #1, Tab(50); "|";
                    Print #1, Tab(51); Pusing("######", xLote(WCounter, 5));
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Pusing("###,###", xLote(WCounter, 6));
                    Print #1, Tab(67); "|";
                    
                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, 5) + "'"
                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEnvase.RecordCount > 0 Then
                        WAbre = rstEnvase!Abreviatura
                        rstEnvase.Close
                            Else
                        WAbre = ""
                    End If
        
                    Print #1, Tab(68); Alinea("###", XEnvase(WCounter, 6)) + " " + Left$(WAbre, 8);
                    
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                If Val(xLote(WCounter, 7)) <> 0 Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|";
                    Print #1, Tab(40); "|";
                    Print #1, Tab(50); "|";
                    Print #1, Tab(51); Pusing("######", xLote(WCounter, 7));
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Pusing("###,###", xLote(WCounter, 8));
                    Print #1, Tab(67); "|";
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                If Val(xLote(WCounter, 9)) <> 0 Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|";
                    Print #1, Tab(40); "|";
                    Print #1, Tab(50); "|";
                    Print #1, Tab(51); Pusing("######", xLote(WCounter, 9));
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Pusing("###,###", xLote(WCounter, 10));
                    Print #1, Tab(67); "|";
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                If Trim(WWObserva) <> "" Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|Observ.:";
                    Print #1, WWObserva;
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                If Trim(WWEspecifica) <> "" Then
                    Print #1, Tab(1); "|";
                    Print #1, Tab(16); "|Especif.:";
                    Print #1, WWEspecifica;
                    Print #1, Tab(80); "|"
                    XLinea = XLinea + 1
                End If
                
                Print #1, Tab(1); String$(79, "-")
                XLinea = XLinea + 1
                    
            End If
            
        Next WCounter
        
        Pasa = 0
        For CicloEspecif = 1 To 10
            If WWEspecif(CicloEspecif) <> "" Then
                If Pasa = 0 Then
                    Print #1, Tab(1); "|Especificaciones : ";
                    Pasa = 1
                End If
                Print #1, Tab(25); WWEspecif(CicloEspecif);
                Print #1, Tab(80); "|"
                XLinea = XLinea + 1
            End If
        Next CicloEspecif
        
        For WDa = XLinea To 14
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(40); "|";
            Print #1, Tab(50); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
        
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "Preparó:      ";
        Print #1, "Etiquetó:       ";
        Print #1, "Fraccionó:      ";
        Print #1, "Supervisó:      ";
        Print #1, "Despachó:       "
        
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""

    Next XX
    
    Print #1, Chr$(12)
    
    Close #1

End Sub




Private Sub ImpresionSql()

    Rem On Error GoTo WError
    
    spImprePedIp = "Delete  ImprePedIp"
    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    WObservaciones = Left$(WWObservaciones + Space$(100), 100)
    Select Case WWTipoped
        Case 0
            WTipoPedido = " (Normal)"
        Case 1
            WTipoPedido = " (A fecha)"
        Case 2
            WTipoPedido = " (Fecha Limite)"
        Case 3
            WTipoPedido = " (Urgente)"
        Case 4
            WTipoPedido = " (Retira Cliente)"
        Case 5
            WTipoPedido = " (Muestra)"
        Case Else
    End Select
    
    WVia = ""
    Select Case Via.ListIndex
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
    
    
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        WWPago = rstCliente!Pago1
        WWDirentrega = ""
        
        ZZDirEntrega(1) = rstCliente!DirEntrega
        ZZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
        ZZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
        ZZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
        ZZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))
        
        WWDirentrega = ZZDirEntrega(ZZLugarDirEntrega)
        
        rstCliente.Close
        
        spPago = "ConsultaPago " + "'" + WWPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WWDesPago = rstPago!Nombre
            rstPago.Close
        End If
        
    End If
    
    
    
        
    XLinea = 0
    WCounter = 0
    WRenglon = 0
                    
    For a = 1 To 40
        
        WCounter = WCounter + 1
                
        If Datos(WCounter, 0) <> "" Then
                
            WArticulo = Datos(WCounter, O)
            WDescripcion = Datos(WCounter, 1)
            WCantidad = Val(Datos(WCounter, 2))
            WPrecio = Val(Datos(WCounter, 3))
            WObserva = Datos(WCounter, 8)
            WEspecificaciones = Datos(WCounter, 9)
                
            If WCantidad <> 0 Then
            
                Erase ImpreEnvase
                LugarEnvase = 0
            
                For Cicla = 1 To 6 Step 2
                    If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                        LugarEnvase = LugarEnvase + 1
                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                        If rstEnvase.RecordCount > 0 Then
                            WAbre = rstEnvase!Abreviatura
                            rstEnvase.Close
                                Else
                            WAbre = ""
                        End If
                        ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8)
                    End If
                Next Cicla
                
                WRenglon = WRenglon + 1
                
                Auxi = Pedido.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                ZClave = "1" + Auxi + Auxi1
                ZTipo = "1"
                ZPedido = Pedido.Text
                ZRenglon = Str$(WRenglon)
                ZEmpresa = ""
                ZVersion = WWVersion
                ZCliente = WWCliente
                ZNombre = DesCliente.Caption
                ZFecha = WWFecha
                ZFechaent = WWFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = WWDesPago
                ZEntrega = WWDirentrega
                
                ZObservaciones1 = Left$(WObservaciones, 50)
                ZObservaciones2 = Right$(WObservaciones, 50)
                ZOrden = WWOrdenCpa
                
                ZArticulo = WArticulo
                ZDescripcion = WDescripcion
                ZPrecio = Str$(WPrecio)
                ZCantidad = Str$(WCantidad)
                ZEnvase = ImpreEnvase(1)
                ZLugarLote = 1
                
                ZZLote = ""
                ZZCantiLote = ""
                Select Case ZLugarLote
                    Case 1
                        ZZLote = xLote(WCounter, 1)
                        ZZCantiLote = xLote(WCounter, 2)
                    Case 2
                        ZZLote = xLote(WCounter, 3)
                        ZZCantiLote = xLote(WCounter, 4)
                    Case 3
                        ZZLote = xLote(WCounter, 5)
                        ZZCantiLote = xLote(WCounter, 6)
                    Case 4
                        ZZLote = xLote(WCounter, 7)
                        ZZCantiLote = xLote(WCounter, 8)
                    Case 5
                        ZZLote = xLote(WCounter, 9)
                        ZZCantiLote = xLote(WCounter, 10)
                    Case Else
                End Select
                
                spImprePedIp = "INSERT INTO ImprePedIp (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase, Lote, CantiLote )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                For Ciclo = 2 To LugarEnvase
                
                    WRenglon = WRenglon + 1
                
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = ""
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = DesCliente.Caption
                    ZFecha = WWFecha
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = ""
                    ZDescripcion = ""
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ImpreEnvase(Ciclo)
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase, Lote, CantiLote )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                    
                Next Ciclo
                
                If Trim(WEspecificaciones) <> "" And WEspecificaciones <> "0" Then
                
                    WRenglon = WRenglon + 1
                
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = DesCliente.Caption
                    ZFecha = Fecha.Text
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = "Especif.:"
                    ZDescripcion = WEspecificaciones
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase, Lote, CantiLote )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Trim(WObserva) <> "" Then
                
                    WRenglon = WRenglon + 1
                
                    Auxi = Pedido.Text
                    Call Ceros(Auxi, 6)
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    ZClave = "1" + Auxi + Auxi1
                    ZTipo = "1"
                    ZPedido = Pedido.Text
                    ZRenglon = Str$(WRenglon)
                    ZEmpresa = WNombreEmpresa
                    ZVersion = WWVersion
                    ZCliente = WWCliente
                    ZNombre = DesCliente.Caption
                    ZFecha = Fecha.Text
                    ZFechaent = WWFecEntrega
                    ZTipoPedido = WTipoPedido
                    ZCondicion = WWDesPago
                    ZEntrega = WWDirentrega
                    ZObservaciones1 = Left$(WObservaciones, 50)
                    ZObservaciones2 = Right$(WObservaciones, 50)
                    ZOrden = WWOrdenCpa
                    ZArticulo = "Observ.:"
                    ZDescripcion = WObserva
                    ZPrecio = "0"
                    ZCantidad = "0"
                    ZEnvase = ""
                    
                    ZLugarLote = ZLugarLote + 1
                    
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case ZLugarLote
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    spImprePedIp = "INSERT INTO ImprePedIp (" + _
                            "Clave ," + _
                            "Tipo , Pedido ," + _
                            "Renglon , Empresa ," + _
                            "Version , Cliente ," + _
                            "Nombre , Fecha ," + _
                            "Fechaent , TipoPedido ," + _
                            "Condicion , Entrega ," + _
                            "Observaciones1 , Observaciones2 ," + _
                            "Orden , Articulo ," + _
                            "Descripcion , Precio ," + _
                            "Cantidad , Envase, Lote, CantiLote )" + _
                            "Values (" + _
                            "'" + ZClave + "'," + _
                            "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                            "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                            "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                            "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                            "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                            "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                            "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                            "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                            "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                            "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                            "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                            
                    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                For Ciclo = ZLugarLote + 1 To 5
                
                    ZZLote = ""
                    ZZCantiLote = ""
                    Select Case Ciclo
                        Case 1
                            ZZLote = xLote(WCounter, 1)
                            ZZCantiLote = xLote(WCounter, 2)
                        Case 2
                            ZZLote = xLote(WCounter, 3)
                            ZZCantiLote = xLote(WCounter, 4)
                        Case 3
                            ZZLote = xLote(WCounter, 5)
                            ZZCantiLote = xLote(WCounter, 6)
                        Case 4
                            ZZLote = xLote(WCounter, 7)
                            ZZCantiLote = xLote(WCounter, 8)
                        Case 5
                            ZZLote = xLote(WCounter, 9)
                            ZZCantiLote = xLote(WCounter, 10)
                        Case Else
                    End Select
                    
                    If Val(ZZLote) <> 0 Or Val(ZZCantiLote) Then
                
                        WRenglon = WRenglon + 1
                    
                        Auxi = Pedido.Text
                        Call Ceros(Auxi, 6)
                        Auxi1 = WRenglon
                        Call Ceros(Auxi1, 2)
                        ZClave = "1" + Auxi + Auxi1
                        ZTipo = "1"
                        ZPedido = Pedido.Text
                        ZRenglon = Str$(WRenglon)
                        ZEmpresa = ""
                        ZVersion = WWVersion
                        ZCliente = WWCliente
                        ZNombre = DesCliente.Caption
                        ZFecha = WWFecha
                        ZFechaent = WWFecEntrega
                        ZTipoPedido = WTipoPedido
                        ZCondicion = WWDesPago
                        ZEntrega = WWDirentrega
                        ZObservaciones1 = Left$(WObservaciones, 50)
                        ZObservaciones2 = Right$(WObservaciones, 50)
                        ZOrden = WWOrdenCpa
                        ZArticulo = ""
                        ZDescripcion = ""
                        ZPrecio = "0"
                        ZCantidad = "0"
                        ZEnvase = ""
                        
                        spImprePedIp = "INSERT INTO ImprePedIp (" + _
                                "Clave ," + _
                                "Tipo , Pedido ," + _
                                "Renglon , Empresa ," + _
                                "Version , Cliente ," + _
                                "Nombre , Fecha ," + _
                                "Fechaent , TipoPedido ," + _
                                "Condicion , Entrega ," + _
                                "Observaciones1 , Observaciones2 ," + _
                                "Orden , Articulo ," + _
                                "Descripcion , Precio ," + _
                                "Cantidad , Envase, Lote, CantiLote )" + _
                                "Values (" + _
                                "'" + ZClave + "'," + _
                                "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                                "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                                "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                                "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                                "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                                "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                                "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                                "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                                "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                                "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                                "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                                
                        Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                Next Ciclo
                    
            End If
                
        End If
            
    Next a
    
    For Ciclo = WRenglon + 1 To 12
    
        WRenglon = WRenglon + 1
        SumaEspe = SumaEspe + 1
    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        ZClave = "1" + Auxi + Auxi1
        ZTipo = "1"
        ZPedido = Pedido.Text
        ZRenglon = Str$(WRenglon)
        ZEmpresa = ""
        ZVersion = WWVersion
        ZCliente = WWCliente
        ZNombre = DesCliente.Caption
        ZFecha = WWFecha
        ZFechaent = WWFecEntrega
        ZTipoPedido = WTipoPedido
        ZCondicion = WWDesPago
        ZEntrega = WWDirentrega
        ZObservaciones1 = Left$(WObservaciones, 50)
        ZObservaciones2 = Right$(WObservaciones, 50)
        ZOrden = WWOrdenCpa
        ZArticulo = ""
        ZDescripcion = ""
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
        ZZLote = ""
        ZZCantiLote = ""
                        
        spImprePedIp = "INSERT INTO ImprePedIp (" + _
                    "Clave ," + _
                    "Tipo , Pedido ," + _
                    "Renglon , Empresa ," + _
                    "Version , Cliente ," + _
                    "Nombre , Fecha ," + _
                    "Fechaent , TipoPedido ," + _
                    "Condicion , Entrega ," + _
                    "Observaciones1 , Observaciones2 ," + _
                    "Orden , Articulo ," + _
                    "Descripcion , Precio ," + _
                    "Cantidad , Envase, Lote, CantiLote )" + _
                    "Values (" + _
                    "'" + ZClave + "'," + _
                    "'" + ZTipo + "'," + "'" + ZPedido + "'," + _
                    "'" + ZRenglon + "'," + "'" + ZEmpresa + "'," + _
                    "'" + ZVersion + "'," + "'" + ZCliente + "'," + _
                    "'" + ZNombre + "'," + "'" + ZFecha + "'," + _
                    "'" + ZFechaent + "'," + "'" + ZTipoPedido + "'," + _
                    "'" + ZCondicion + "'," + "'" + ZEntrega + "'," + _
                    "'" + ZObservaciones1 + "'," + "'" + ZObservaciones2 + "'," + _
                    "'" + ZOrden + "'," + "'" + ZArticulo + "'," + _
                    "'" + ZDescripcion + "'," + "'" + ZPrecio + "'," + _
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "'," + _
                    "'" + ZZLote + "'," + "'" + ZZCantiLote + "')"
                                
        Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePedIp SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "',"
    ZSql = ZSql + "TipoPed = " + "'" + WWTipoped + "'"
    spImprePedIp = ZSql
    Set rstImprePedIp = db.OpenRecordset(spImprePedIp, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePedIp ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    Listado.ReportFileName = "ImprePedsqlip.rpt"
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    
    
    
    ZZRequiereCertificado = ""
    ZZRequiereMsds = ""
    ZZRequiereMsdsCada = ""
    ZZRequiereHoja = ""
    ZZPermiteParcial = ""
    ZZPartidasVarias = ""

    ZZEmailCertificado = ""
    ZZEmailMsds = ""
    ZZEmailHoja = ""
    ZZDiasI = ""
    ZZDiasII = ""
    ZZDiasIII = ""
    ZZEnvasesI = ""
    ZZEnvasesII = ""
    ZZEnvasesIII = ""
    ZZEtiquetaI = ""
    ZZEtiquetaII = ""
    ZZEspecif1 = ""
    ZZEspecif2 = ""
    ZZEspecif3 = ""
    ZZEspecif4 = ""
    ZZEspecif5 = ""
    ZZCantidadPartidas = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteEspecif"
    ZSql = ZSql + " Where ClienteEspecif.Cliente = " + "'" + WWCliente + "'"
    spClienteEspecif = ZSql
    Set rstClienteEspecif = db.OpenRecordset(spClienteEspecif, dbOpenSnapshot, dbSQLPassThrough)
    If rstClienteEspecif.RecordCount > 0 Then
    
        ZZRequiereCertificado = IIf(IsNull(rstClienteEspecif!RequiereCertificado), "0", rstClienteEspecif!RequiereCertificado)
        ZZRequiereMsds = IIf(IsNull(rstClienteEspecif!RequiereMsds), "0", rstClienteEspecif!RequiereMsds)
        ZZRequiereMsdsCada = IIf(IsNull(rstClienteEspecif!RequiereMsdsCada), "0", rstClienteEspecif!RequiereMsdsCada)
        ZZRequiereHoja = IIf(IsNull(rstClienteEspecif!RequiereHoja), "0", rstClienteEspecif!RequiereHoja)
        ZZPermiteParcial = IIf(IsNull(rstClienteEspecif!PermiteParcial), "0", rstClienteEspecif!PermiteParcial)
        ZZPartidasVarias = IIf(IsNull(rstClienteEspecif!PartidaVarias), "0", rstClienteEspecif!PartidaVarias)

        ZZEmailCertificado = IIf(IsNull(rstClienteEspecif!EmailCertificado), "", rstClienteEspecif!EmailCertificado)
        ZZEmailMsds = IIf(IsNull(rstClienteEspecif!EmailMsds), "", rstClienteEspecif!EmailMsds)
        ZZEmailHoja = IIf(IsNull(rstClienteEspecif!EmailHoja), "", rstClienteEspecif!EmailHoja)
        ZZDiasI = IIf(IsNull(rstClienteEspecif!DiasI), "", rstClienteEspecif!DiasI)
        ZZDiasII = IIf(IsNull(rstClienteEspecif!DiasII), "", rstClienteEspecif!DiasII)
        ZZDiasIII = IIf(IsNull(rstClienteEspecif!DiasIII), "", rstClienteEspecif!DiasIII)
        ZZEnvasesI = IIf(IsNull(rstClienteEspecif!EnvasesI), "", rstClienteEspecif!EnvasesI)
        ZZEnvasesII = IIf(IsNull(rstClienteEspecif!EnvasesII), "", rstClienteEspecif!EnvasesII)
        ZZEnvasesIII = IIf(IsNull(rstClienteEspecif!EnvasesIII), "", rstClienteEspecif!EnvasesIII)
        ZZEtiquetaI = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
        ZZEtiquetaII = IIf(IsNull(rstClienteEspecif!EtiquetaI), "", rstClienteEspecif!EtiquetaI)
        ZZEspecif1 = IIf(IsNull(rstClienteEspecif!Especif1), "", rstClienteEspecif!Especif1)
        ZZEspecif2 = IIf(IsNull(rstClienteEspecif!Especif2), "", rstClienteEspecif!Especif2)
        ZZEspecif3 = IIf(IsNull(rstClienteEspecif!Especif3), "", rstClienteEspecif!Especif3)
        ZZEspecif4 = IIf(IsNull(rstClienteEspecif!Especif4), "", rstClienteEspecif!Especif4)
        ZZEspecif5 = IIf(IsNull(rstClienteEspecif!Especif5), "", rstClienteEspecif!Especif5)
        ZZCantidadPartidas = IIf(IsNull(rstClienteEspecif!CantidadPartidas), "", rstClienteEspecif!CantidadPartidas)
        
        rstClienteEspecif.Close
        
    End If
    
    ZZImprime = "N"
    
    If Val(ZZRequiereCertificado) <> 0 Or Val(ZZRequiereMsds) <> 0 Or Val(ZZRequiereMsdsCada) <> 0 Or Val(ZZRequiereHoja) <> 0 Or Val(ZZPermiteParcial) <> 0 Or Val(ZZPartidasVarias) <> 0 Then
        ZZImprime = "S"
    End If
    If Trim(ZZDiasI) <> "" Or Trim(ZZDiasII) <> "" Or Trim(ZZDiasIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEnvasesI) <> "" Or Trim(ZZEnvasesII) <> "" Or Trim(ZZEnvasesIII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEtiquetaI) <> "" Or Trim(ZZEtiquetaII) <> "" Then
        ZZImprime = "S"
    End If
    If Trim(ZZEspecif1) <> "" Or Trim(ZZEspecif2) <> "" Or Trim(ZZEspecif3) <> "" Or Trim(ZZEspecif4) <> "" Or Trim(ZZEspecif5) <> "" Then
        ZZImprime = "S"
    End If
    If Val(ZZCantidadPartidas) <> 0 Then
        ZZImprime = "S"
    End If
    
    If ZZImprime = "S" Then
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT ImprePed.Clave, ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.TipoPedido, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via, " _
                + "ClienteEspecif.RequiereCertificado, ClienteEspecif.RequiereMsds, ClienteEspecif.RequiereMsdsCada, ClienteEspecif.RequiereHoja, ClienteEspecif.PermiteParcial, ClienteEspecif.DiasI, ClienteEspecif.DiasII, ClienteEspecif.DiasIII, ClienteEspecif.Especif1, ClienteEspecif.Especif2, ClienteEspecif.Especif3, ClienteEspecif.Especif4, ClienteEspecif.Especif5, ClienteEspecif.PartidaVarias, ClienteEspecif.CantidadPartidas, ClienteEspecif.EnvasesI, ClienteEspecif.EnvasesII, ClienteEspecif.EnvasesIII, ClienteEspecif.EtiquetaI, ClienteEspecif.EtiquetaII " _
                + "From " _
                + DSQ + ".dbo.ImprePedIp ImprePed, " _
                + DSQ + ".dbo.ClienteEspecif ClienteEspecif " _
                + "Where " _
                + "ImprePed.Cliente = ClienteEspecif.Cliente AND " _
                + "ImprePed.Pedido >= 0 AND " _
                + "ImprePed.Pedido <= 999999"
                            
        Listado.Connect = Connect()
        Listado.ReportFileName = "ImprePedsqlEspecifIp.rpt"
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.CopiesToPrinter = 1
        Listado.Action = 1
        
    End If
        
    Exit Sub
        
WError:
    Resume Next

End Sub



Private Sub Verifica_Restriccion()

    If Val(WEmpresa) = 1 Then
    
        ZZRestriccion = 0
        ZZZPasa = "S"
        
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZRestriccion = IIf(IsNull(rstCliente!Restriccion), "0", rstCliente!Restriccion)
            rstCliente.Close
        End If
        
        If ZZRestriccion = 1 Then
        
            If Left$(ZZZTerminado, 2) <> "PT" And Left$(ZZZTerminado, 2) <> "YQ" And Left$(ZZZTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            ZZRestriccionI = 0
            ZZRestriccionII = 0
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(ZZZTerminado, 3) + Right$(ZZZTerminado, 7)
                    
                    spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        ZZRestriccionII = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                        rstArticulo.Close
                    End If
                    If ZZRestriccionII = 1 Then
                        ZZZPasa = "N"
                        m$ = "El cliente posee restriccion para los productos" + Chr$(13) + _
                             "y algun componente de esta partida lo posee"
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    End If
    
                Case Else
                    XEmpresa = WEmpresa
                    
                    ZZLugarVeri = 0
                    Erase ZZVerifica
                    
                                
                    Select Case Val(WEmpresa)
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
                            ZHasta = 7
                        Case Else
                            CargaEmpresa(1, 1) = "0002"
                            CargaEmpresa(1, 2) = "Empresa02"
                            CargaEmpresa(2, 1) = "0004"
                            CargaEmpresa(2, 2) = "Empresa04"
                            CargaEmpresa(3, 1) = "0008"
                            CargaEmpresa(3, 2) = "Empresa08"
                            CargaEmpresa(4, 1) = "0009"
                            CargaEmpresa(4, 2) = "Empresa09"
                            ZHasta = 4
                    End Select
                    
                    For CiclaEmpresa = 1 To ZHasta
                    
                        WEmpresa = CargaEmpresa(CiclaEmpresa, 1)
                        txtOdbc = CargaEmpresa(CiclaEmpresa, 2)
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                        spHoja = "ListaHoja " + "'" + ZZZLote + "'"
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            With rstHoja
                                .MoveFirst
                                Do
                                    If .EOF = False Then
                                        ZZLugarVeri = ZZLugarVeri + 1
                                        ZZVerifica(ZZLugarVeri, 1) = rstHoja!Tipo
                                        If UCase(rstHoja!Tipo) = "M" Then
                                            ZZVerifica(ZZLugarVeri, 2) = rstHoja!Articulo
                                                Else
                                            ZZVerifica(ZZLugarVeri, 2) = rstHoja!Terminado
                                        End If
                                        .MoveNext
                                            Else
                                        Exit Do
                                    End If
                                Loop
                            End With
                            rstHoja.Close
                        End If
                    
                    Next CiclaEmpresa
                    
                    For CicloVeri = 1 To ZZLugarVeri
                    
                        ZZTipoVeri = ZZVerifica(CicloVeri, 1)
                        
                        If UCase(ZZTipoVeri) = "M" Then
                        
                            ZZArtiVeri = ZZVerifica(CicloVeri, 2)
                            
                            spArticulo = "ConsultaArticulo " + "'" + ZZArtiVeri + "'"
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                ZZRestriccionI = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                                rstArticulo.Close
                            End If
                            If ZZRestriccionI = 1 Then
                                ZZRestriccionII = 1
                            End If
                            
                                Else
                                
                            ZZTermiVeri = ZZVerifica(CicloVeri, 2)
                                                    
                            spTerminado = "ConsultaTerminado " + "'" + ZZTermiVeri + "'"
                            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            If rstTerminado.RecordCount > 0 Then
                                ZZRestriccionI = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
                                rstTerminado.Close
                            End If
                            If ZZRestriccionI = 1 Then
                                ZZRestriccionII = 1
                            End If
                            
                        End If
                        
                    Next CicloVeri
                    
                    Call Conecta_Empresa
                            
                    If ZZRestriccionII = 1 Then
                        ZZZPasa = "N"
                        m$ = "El cliente posee restriccion para los productos" + Chr$(13) + _
                             "y algun componente de esta partida lo posee"
                        G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                    End If
            
            End Select
                    
        End If
        
    End If


End Sub
        
        
        
        


