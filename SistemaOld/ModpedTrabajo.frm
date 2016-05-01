VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModped 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Actualizacion de Pedidos a Facturar"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   11550
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   11550
   Visible         =   0   'False
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
      Height          =   2655
      Left            =   2760
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox WEnvase5 
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
         TabIndex        =   61
         Text            =   " "
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox WCantiEnv5 
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
         Left            =   5280
         TabIndex        =   60
         Text            =   " "
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox WEnvase4 
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
         TabIndex        =   58
         Text            =   " "
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox WCantiEnv4 
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
         Left            =   5280
         TabIndex        =   57
         Text            =   " "
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox WEnvase3 
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
         TabIndex        =   55
         Text            =   " "
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox WCantiEnv3 
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
         Left            =   5280
         TabIndex        =   54
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WEnvase2 
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
         TabIndex        =   52
         Text            =   " "
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox WCantiEnv2 
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
         Left            =   5280
         TabIndex        =   51
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WEnvase1 
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
         TabIndex        =   46
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox WCantiEnv1 
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
         Left            =   5280
         TabIndex        =   45
         Text            =   " "
         Top             =   600
         Width           =   855
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
         Left            =   1800
         TabIndex        =   42
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
         Left            =   1800
         TabIndex        =   41
         Top             =   1680
         Width           =   975
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1680
         Width           =   1335
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
         Left            =   1800
         TabIndex        =   38
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
         Left            =   1800
         TabIndex        =   37
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
         Left            =   1800
         TabIndex        =   36
         Top             =   600
         Width           =   975
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   34
         Top             =   960
         Width           =   1335
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label WDescri5 
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
         Left            =   4200
         TabIndex        =   62
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label WDescri4 
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
         Left            =   4200
         TabIndex        =   59
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label WDescri3 
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
         Left            =   4200
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label WDescri2 
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
         Left            =   4200
         TabIndex        =   53
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Canti."
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
         Left            =   5280
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   4200
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Envase"
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
         Left            =   3000
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.Label WDescri1 
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
         Left            =   4200
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
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
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   975
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
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
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
      Height          =   495
      Left            =   8640
      TabIndex        =   44
      Top             =   720
      Width           =   1335
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
      Height          =   500
      Left            =   7200
      TabIndex        =   29
      Top             =   720
      Width           =   1335
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
      TabIndex        =   23
      Text            =   " "
      Top             =   7680
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
      TabIndex        =   22
      Text            =   " "
      Top             =   7320
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
      TabIndex        =   21
      Text            =   " "
      Top             =   6960
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
      TabIndex        =   20
      Text            =   " "
      Top             =   6600
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
      TabIndex        =   19
      Text            =   " "
      Top             =   6240
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
      TabIndex        =   18
      Text            =   " "
      Top             =   7680
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
      TabIndex        =   17
      Text            =   " "
      Top             =   7320
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
      TabIndex        =   16
      Text            =   " "
      Top             =   6960
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
      TabIndex        =   15
      Text            =   " "
      Top             =   6600
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
      TabIndex        =   14
      Text            =   " "
      Top             =   6240
      Width           =   975
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
      TabIndex        =   13
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
      Left            =   10080
      TabIndex        =   11
      Top             =   120
      Width           =   1335
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
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
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
      TabIndex        =   8
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
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
      Left            =   7200
      TabIndex        =   4
      Top             =   120
      Width           =   1335
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
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   0
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
      Height          =   2220
      ItemData        =   "ModpedTrabajo.frx":0000
      Left            =   3360
      List            =   "ModpedTrabajo.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   8055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "ModpedTrabajo.frx":0015
      TabIndex        =   2
      Top             =   1320
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ENVASES A ENTREGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   5880
      Width           =   3015
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
      TabIndex        =   28
      Top             =   7680
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
      TabIndex        =   27
      Top             =   7320
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
      TabIndex        =   26
      Top             =   6960
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
      TabIndex        =   25
      Top             =   6600
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
      TabIndex        =   24
      Top             =   6240
      Width           =   855
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
      TabIndex        =   12
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
      TabIndex        =   9
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
      TabIndex        =   7
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
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgModped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WAceptada As String
Private WDirentrega As String
Private WFecEntrega As String
Private WDespago As String
Private WObservaciones As String
Private WVersion As String
Private WTipoPedido As String

Private WImpre(10) As String
Private WEnvase(10) As String
Private Envase(5, 2) As String
Private Auxiliar(100, 14) As String
Private ClavePedido(100) As String
Private BajaLote(5, 2) As String
Private xLote(100, 30) As String
Private XEnvase(100, 6) As String
Private EmiteCerti(1000, 3) As String
Private CargaEmpresa(10, 2) As String
Private LugarCerti As Integer

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
Dim rstEnvase As Recordset
Dim spEnvase As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstPago As Recordset
Dim spPago As String

Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstCertificado As Recordset
Dim spCertificado As String

Dim ZOpcion(10) As Integer
Dim ZValor(10) As String
Dim ZEnsayo(10) As String
Dim ZStd(10, 2) As String
Dim ZDescri(10) As String
Dim ZClave1 As String
Dim ZClave2 As String

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
Dim XMes As String
Dim XAno As String

Dim ControlLote(5, 2) As String
Dim WSaldo As Double
Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer

Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String

Private Sub Borra_Click()

    Rem DBGrid1.Col = 0
    Rem DBGrid1.Text = ""
    
    Rem DBGrid1.Col = 1
    Rem DBGrid1.Text = ""

    Rem DBGrid1.Col = 2
    Rem DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 5
    DBGrid1.Text = "S"
    
    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
    
    xLote(WLugar, 1) = ""
    xLote(WLugar, 2) = ""
    xLote(WLugar, 3) = ""
    xLote(WLugar, 4) = ""
    xLote(WLugar, 5) = ""
    xLote(WLugar, 6) = ""
    xLote(WLugar, 7) = ""
    xLote(WLugar, 8) = ""
    xLote(WLugar, 9) = ""
    xLote(WLugar, 10) = ""
    
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstEmpresa
        .Close
    End With
    
    PrgModped.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    WRenglon = 0
    DBGrid1.Refresh
        
    For A = 0 To 7
    
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            
            WRenglon = WRenglon + 1
            WRow = iRow
            DBGrid1.Row = WRow
                
            DBGrid1.Col = 3
            Cantidad = Val(DBGrid1.Text)
                
            DBGrid1.Col = 4
            Resta = Val(DBGrid1.Text)
                
            If Cantidad <> 0 Or Resta <> 0 Then
                DBGrid1.Col = 5
                If DBGrid1.Text <> "S" Then
                    m$ = "No asigno las partidas a todos los productos"
                    A = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    DBGrid1.Refresh
                    Exit Sub
                End If
            End If
               
        Next iRow
            
    Next A
    
    SumaEnvases = Val(Canti1.Text) + Val(Canti2.Text) + Val(Canti3.Text) + Val(Canti4.Text) + Val(Canti5.Text)
    If SumaEnvases = 0 Then
    
        T$ = "Actualizacion de Datos del Pedido - Entrega de Envases"
        m$ = "NO SE INFORMO NINGUN ENVASE A ENTREGAR" + Chr$(13) + "AL CLIENTE  EN EL PRESENTE ENVIO" + Chr$(13) + "" + Chr$(13) + "CONFIRMA LA GRABACION DE LOS DATOS ?"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 7 Then
            DBGrid1.FirstRow = 0
            DBGrid1.Row = 0
            DBGrid1.Col = 0
            Exit Sub
        End If
            Else
        T$ = "Actualizacion de Datos del Pedido - Entrega de Envases"
        m$ = "SE INFORMARON LOS SIGUIENTES ENVASES A ENVIAR AL CLIENTE" + Chr$(13) + "" + Chr$(13) + ""
        If Val(Canti1.Text) <> 0 Then
            m$ = m$ + Canti1.Text + " envases de " + Descri1.Caption + Chr$(13) + ""
        End If
        If Val(Canti2.Text) <> 0 Then
            m$ = m$ + Canti2.Text + " envases de " + descri2.Caption + Chr$(13) + ""
        End If
        If Val(Canti3.Text) <> 0 Then
            m$ = m$ + Canti3.Text + " envases de " + Descri3.Caption + Chr$(13) + ""
        End If
        If Val(Canti4.Text) <> 0 Then
            m$ = m$ + Canti4.Text + " envases de " + Descri4.Caption + Chr$(13) + ""
        End If
        If Val(Canti5.Text) <> 0 Then
            m$ = m$ + Canti5.Text + " envases de " + Descri5.Caption + Chr$(13) + ""
        End If
        m$ = m$ + "" + Chr$(13) + "CONFIRMA LA GRABACION DE LOS DATOS ?"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 7 Then
            DBGrid1.FirstRow = 0
            DBGrid1.Row = 0
            DBGrid1.Col = 0
            Exit Sub
        End If
        
    End If
        
    Erase Auxiliar
    Auxi = 0
        
    Suma = 0
    Renglon = 0
    WRenglon = 0
    DBGrid1.Refresh
    LugarCerti = 0
    Erase EmiteCerti
        
    For A = 0 To 7
        
        Suma = A * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
            
            Suma = Suma + 1
            WRenglon = WRenglon + 1
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 0
            Articulo = DBGrid1.Text
                    
            DBGrid1.Col = 3
            Cantidad = DBGrid1.Text
                    
            DBGrid1.Col = 4
            Resta = Val(DBGrid1.Text)
                    
            Auxi = Pedido.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            XPedido = Left$(ClavePedido(WRenglon), 6)
            XRenglon = Right$(ClavePedido(WRenglon), 2)
            
            WClavePedido = ClavePedido(WRenglon)
            
            If Trim(Articulo) <> "" Then
                
                XCantidad1 = Cantidad
                xCantidad2 = Cantidad
                    
                WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    
                XLote1 = xLote(WLugar, 1)
                XLote2 = xLote(WLugar, 3)
                XLote3 = xLote(WLugar, 5)
                XLote4 = xLote(WLugar, 7)
                XLote5 = xLote(WLugar, 9)
                    
                If Left$(Articulo, 2) <> "PT" Then
                
                    For Ciclo = 1 To 9 Step 2
                    
                        If xLote(WLugar, Ciclo) <> "" Then
                            
                            ZEntra = "N"
                            
                            XEmpresa = WEmpresa
                            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                                Select Case WTipoPedido
                                    Case "PG", "CO"
                                        WEmpresa = "0001"
                                        txtOdbc = "Empresa01"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case "FA"
                                        WEmpresa = "0005"
                                        txtOdbc = "Empresa05"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                    Case Else
                                        WEmpresa = "0007"
                                        txtOdbc = "Empresa07"
                                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                End Select
                            End If
                                
                            Sql1 = "Select *"
                            Sql2 = " FROM Laudo"
                            Sql3 = " Where Laudo.PartiOri = " + "'" + xLote(WLugar, Ciclo) + "'"
                            Sql4 = " Order by Laudo.Laudo"
                            spLaudo = Sql1 + Sql2 + Sql3 + Sql4
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                With rstLaudo
                                    .MoveFirst
                                    Select Case Ciclo
                                        Case 1
                                            XLote1 = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                                        Case 3
                                            XLote2 = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                                        Case 5
                                            XLote3 = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                                        Case 7
                                            XLote4 = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                                        Case 9
                                            XLote5 = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
                                        Case Else
                                    End Select
                                    ZEntra = "S"
                                    rstLaudo.Close
                                End With
                            End If
                        
                            If ZEntra = "N" Then
                                Sql1 = "Select *"
                                Sql2 = " FROM Guia"
                                Sql3 = " Where Guia.PartiOri = " + "'" + xLote(WLugar, Ciclo) + "'"
                                Sql4 = " Order by Guia.Saldo desc"
                                spMovguia = Sql1 + Sql2 + Sql3 + Sql4
                                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                If rstMovguia.RecordCount > 0 Then
                                    With rstMovguia
                                        .MoveFirst
                                        Select Case Ciclo
                                            Case 1
                                                XLote1 = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                            Case 3
                                                XLote2 = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                            Case 5
                                                XLote3 = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                            Case 7
                                                XLote4 = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                            Case 9
                                                XLote5 = IIf(IsNull(rstMovguia!Lote), "", rstMovguia!Lote)
                                            Case Else
                                        End Select
                                        ZEntra = "S"
                                        rstMovguia.Close
                                    End With
                                End If
                            End If
                            
                            Call Conecta_Empresa
                            
                        End If
                            
                    Next Ciclo
                    
                        Else
                                            
                    XLote1 = xLote(WLugar, 1)
                    XLote2 = xLote(WLugar, 3)
                    XLote3 = xLote(WLugar, 5)
                    XLote4 = xLote(WLugar, 7)
                    XLote5 = xLote(WLugar, 9)
                
                End If
                    
                XCantiLote1 = xLote(WLugar, 2)
                XCantiLote2 = xLote(WLugar, 4)
                XCantiLote3 = xLote(WLugar, 6)
                XCantiLote4 = xLote(WLugar, 8)
                XCantiLote5 = xLote(WLugar, 10)
                XEnv1 = Envase1.Text
                XCantiEnv1 = Canti1.Text
                XEnv2 = Envase2.Text
                XCantiEnv2 = Canti2.Text
                XEnv3 = Envase3.Text
                XCantiEnv3 = Canti3.Text
                XEnv4 = Envase4.Text
                XCantiEnv4 = Canti4.Text
                XEnv5 = Envase5.Text
                XCantiEnv5 = Canti5.Text
                XEti1 = ""
                XEti2 = ""
                XEti3 = ""
                XEti4 = ""
                XEti5 = ""
                XTipo1 = ""
                XTipo2 = ""
                XTipo3 = ""
                XTipo4 = ""
                XTipo5 = ""
                
                ZFechaActualiza = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZOrdFechaActualiza = Right$(ZFechaActualiza, 4) + Mid$(ZFechaActualiza, 4, 2) + Left$(ZFechaActualiza, 2)
                
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Pedido SET "
                ZSql = ZSql & "Cantidad1 = " + "'" + XCantidad1 + "',"
                ZSql = ZSql & "Cantidad2 = " + "'" + xCantidad2 + "',"
                ZSql = ZSql & "Lote1 = " + "'" + XLote1 + "',"
                ZSql = ZSql & "CantiLote1 = " + "'" + XCantiLote1 + "',"
                ZSql = ZSql & "Lote2 = " + "'" + XLote2 + "',"
                ZSql = ZSql & "CantiLote2 = " + "'" + XCantiLote2 + "',"
                ZSql = ZSql & "Lote3 = " + "'" + XLote3 + "',"
                ZSql = ZSql & "CantiLote3 = " + "'" + XCantiLote3 + "',"
                ZSql = ZSql & "Lote4 = " + "'" + XLote4 + "',"
                ZSql = ZSql & "CantiLote4 = " + "'" + XCantiLote4 + "',"
                ZSql = ZSql & "Lote5 = " + "'" + XLote5 + "',"
                ZSql = ZSql & "CantiLote5 = " + "'" + XCantiLote5 + "',"
                ZSql = ZSql & "Env1 = " + "'" + XEnv1 + "',"
                ZSql = ZSql & "CantiEnv1 = " + "'" + XCantiEnv1 + "',"
                ZSql = ZSql & "Env2 = " + "'" + XEnv2 + "',"
                ZSql = ZSql & "CantiEnv2 = " + "'" + XCantiEnv2 + "',"
                ZSql = ZSql & "Env3 = " + "'" + XEnv3 + "',"
                ZSql = ZSql & "CantiEnv3 = " + "'" + XCantiEnv3 + "',"
                ZSql = ZSql & "Env4 = " + "'" + XEnv4 + "',"
                ZSql = ZSql & "CantiEnv4 = " + "'" + XCantiEnv4 + "',"
                ZSql = ZSql & "Env5 = " + "'" + XEnv5 + "',"
                ZSql = ZSql & "CantiEnv5 = " + "'" + XCantiEnv5 + "',"
                ZSql = ZSql & "FechaActualizacion = " + "'" + ZFechaActualiza + "',"
                ZSql = ZSql & "OrdFechaActualizacion = " + "'" + ZOrdFechaActualiza + "'"
                ZSql = ZSql & " Where Clave = " + "'" + WClavePedido + "'"
                
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                Call Conecta_Empresa
    
                If Val(XLote1) <> 0 Then
                    LugarCerti = LugarCerti + 1
                    EmiteCerti(LugarCerti, 1) = Articulo
                    EmiteCerti(LugarCerti, 2) = XLote1
                    EmiteCerti(LugarCerti, 3) = XCantiLote1
                End If
                
                If Val(XLote2) <> 0 Then
                    LugarCerti = LugarCerti + 1
                    EmiteCerti(LugarCerti, 1) = Articulo
                    EmiteCerti(LugarCerti, 2) = XLote2
                    EmiteCerti(LugarCerti, 3) = XCantiLote2
                End If
                
                If Val(XLote3) <> 0 Then
                    LugarCerti = LugarCerti + 1
                    EmiteCerti(LugarCerti, 1) = Articulo
                    EmiteCerti(LugarCerti, 2) = XLote3
                    EmiteCerti(LugarCerti, 3) = XCantiLote3
                End If
                
                If Val(XLote4) <> 0 Then
                    LugarCerti = LugarCerti + 1
                    EmiteCerti(LugarCerti, 1) = Articulo
                    EmiteCerti(LugarCerti, 2) = XLote4
                    EmiteCerti(LugarCerti, 3) = XCantiLote4
                End If
                    
                If Val(XLote5) <> 0 Then
                    LugarCerti = LugarCerti + 1
                    EmiteCerti(LugarCerti, 1) = Articulo
                    EmiteCerti(LugarCerti, 2) = XLote5
                    EmiteCerti(LugarCerti, 3) = XCantiLote5
                End If
                
            End If
                                        
        Next iRow
            
    Next A
    
    Rem Call Imprime_Certificado
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Exit Sub

WError:
     Resume Next
        
End Sub

Private Sub Limpia_Click()

    Erase XEnvase

    CargaLote.Visible = False
    Erase xLote
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    
    For A = 0 To 7
        Suma = A * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next A
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Envase1.Text = ""
    Envase2.Text = ""
    Envase3.Text = ""
    Envase4.Text = ""
    Envase5.Text = ""
    
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    
    Canti1.Text = ""
    Canti2.Text = ""
    Canti3.Text = ""
    Canti4.Text = ""
    Canti5.Text = ""
    
    Pedido.SetFocus

End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 3
                Select Case KeyCode
                    Case 13
                        DBGrid1.Col = 3
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 4
                        KeyCode = 0
                    Case Else
                End Select
                        
            Case 4
                Select Case KeyCode
                    Case 13
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Val(DBGrid1.Text)))
                        DBGrid1.Col = 0
                        XTerminado = DBGrid1.Text
                        DBGrid1.Col = 3
                        XCantidad = Val(DBGrid1.Text)
                        WRow = DBGrid1.Row
                        
                        CargaLote.Visible = True
                        WLote1.Text = ""
                        WCanti1.Text = ""
                        WLote2.Text = ""
                        WCanti2.Text = ""
                        WLote3.Text = ""
                        WCanti3.Text = ""
                        WLote4.Text = ""
                        WCanti4.Text = ""
                        WLote5.Text = ""
                        WCanti5.Text = ""
                        
                        WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                            
                        If xLote(WLugar, 1) <> "" Then
                            WLote1.Text = xLote(WLugar, 1)
                            WCanti1.Text = xLote(WLugar, 2)
                        End If
                        If xLote(WLugar, 3) <> "" Then
                            WLote2.Text = xLote(WLugar, 3)
                            WCanti2.Text = xLote(WLugar, 4)
                        End If
                        If xLote(WLugar, 5) <> "" Then
                            WLote3.Text = xLote(WLugar, 5)
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
                            
                        WLote1.SetFocus
                        
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote1.Text + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WEntra = "S"
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo1 = 0
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote1.Text + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
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
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                    
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
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
                        
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                
                    End If
            End Select
                
            If WLote1.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    xLote(WLugar, 1) = WLote1.Text
                    xLote(WLugar, 2) = WCanti1.Text
                    xLote(WLugar, 3) = WLote2.Text
                    xLote(WLugar, 4) = WCanti2.Text
                    xLote(WLugar, 5) = WLote3.Text
                    xLote(WLugar, 6) = WCanti3.Text
                    xLote(WLugar, 7) = WLote4.Text
                    xLote(WLugar, 8) = WCanti4.Text
                    xLote(WLugar, 9) = WLote5.Text
                    xLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                        DBGrid1.Row = DBGrid1.Row + 1
                        WRow = DBGrid1.Row
                        XRow = DBGrid1.Row
                        DBGrid1.Col = 3
                        KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote1.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti1.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente"
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            End If
            
            If WEstado = "N" Then
                WLote1.SetFocus
            End If
            
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WEnvase1.SetFocus
                Else
            XSaldo1 = WSaldo1
            XSaldo1 = Pusing("###,###.##", XSaldo1)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo1
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            WLote1.SetFocus
        End If
        Rem WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
        Rem WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEnvase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + WEnvase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            WDescri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            WCantiEnv1.SetFocus
                Else
            WEnvase1.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantiEnv1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote2.Text + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WEntra = "S"
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo2 = 0
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote2.Text + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
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
                            End With
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
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
                        
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
            End Select
                
            If WLote2.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    xLote(WLugar, 1) = WLote1.Text
                    xLote(WLugar, 2) = WCanti1.Text
                    xLote(WLugar, 3) = WLote2.Text
                    xLote(WLugar, 4) = WCanti2.Text
                    xLote(WLugar, 5) = WLote3.Text
                    xLote(WLugar, 6) = WCanti3.Text
                    xLote(WLugar, 7) = WLote4.Text
                    xLote(WLugar, 8) = WCanti4.Text
                    xLote(WLugar, 9) = WLote5.Text
                    xLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote2.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti2.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente"
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            End If
            
            If WEstado = "N" Then
                WLote2.SetFocus
            End If
    
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WLote3.SetFocus
                Else
            XSaldo2 = WSaldo2
            XSaldo2 = Pusing("###,###.##", XSaldo2)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo2
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            WLote2.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                    
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote3.Text + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WEntra = "S"
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo3 = 0
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote3.Text + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
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
                            End With
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                        
                        XParam = "'" + WLote3.Text + "','" _
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
                                    + WLote3.Text + "'"
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
                        
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
                    
            End Select
                
            If WLote3.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    xLote(WLugar, 1) = WLote1.Text
                    xLote(WLugar, 2) = WCanti1.Text
                    xLote(WLugar, 3) = WLote2.Text
                    xLote(WLugar, 4) = WCanti2.Text
                    xLote(WLugar, 5) = WLote3.Text
                    xLote(WLugar, 6) = WCanti3.Text
                    xLote(WLugar, 7) = WLote4.Text
                    xLote(WLugar, 8) = WCanti4.Text
                    xLote(WLugar, 9) = WLote5.Text
                    xLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote3.Text + " inexistente"
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            End If
            
            If WEstado = "N" Then
                WLote3.SetFocus
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
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            WLote3.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote4.Text + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WEntra = "S"
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo4 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo4 = 0
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote4.Text + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
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
                            End With
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
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
                        
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
                    
            End Select
                
            If WLote4.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    xLote(WLugar, 1) = WLote1.Text
                    xLote(WLugar, 2) = WCanti1.Text
                    xLote(WLugar, 3) = WLote2.Text
                    xLote(WLugar, 4) = WCanti2.Text
                    xLote(WLugar, 5) = WLote3.Text
                    xLote(WLugar, 6) = WCanti3.Text
                    xLote(WLugar, 7) = WLote4.Text
                    xLote(WLugar, 8) = WCanti4.Text
                    xLote(WLugar, 9) = WLote5.Text
                    xLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    Exit Sub
                        Else
                    WLote4.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti4.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente"
                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            End If
            
            If WEstado = "N" Then
                WLote4.SetFocus
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
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            WLote4.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                
                    Sql1 = "Select *"
                    Sql2 = " FROM Laudo"
                    Sql3 = " Where Laudo.Articulo = " + "'" + WArti + "'"
                    Sql4 = " and Laudo.PartiOri = " + "'" + WLote4.Text + "'"
                    Sql5 = " Order by Laudo.Laudo"
                    spLaudo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WEntra = "S"
                            WEstado = IIf(IsNull(rstLaudo!Estado), "", rstLaudo!Estado)
                            If WEstado <> "N" Then
                                WSaldo5 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                                    Else
                                WEstadoII = IIf(IsNull(rstLaudo!EstadoII), "", rstLaudo!EstadoII)
                                If WEstadoII = "V" Then
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por inactividad > a 24 meses)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                        Else
                                    m$ = "La Partida se encuentra bloqueada en espera de la confirmacion de su estado por parte del laboratorio (por devolucion de mercaderia)"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                End If
                                WSaldo5 = 0
                            End If
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                        Sql1 = "Select *"
                        Sql2 = " FROM Guia"
                        Sql3 = " Where Guia.Articulo = " + "'" + WArti + "'"
                        Sql4 = " and Guia.PartiOri = " + "'" + WLote4.Text + "'"
                        Sql5 = " Order by Guia.Saldo desc"
                        spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
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
                            End With
                        End If
                    End If
                    
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
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
                        
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                    
                    End If
                    
            End Select
                
            If WLote5.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    xLote(WLugar, 1) = WLote1.Text
                    xLote(WLugar, 2) = WCanti1.Text
                    xLote(WLugar, 3) = WLote2.Text
                    xLote(WLugar, 4) = WCanti2.Text
                    xLote(WLugar, 5) = WLote3.Text
                    xLote(WLugar, 6) = WCanti3.Text
                    xLote(WLugar, 7) = WLote4.Text
                    xLote(WLugar, 8) = WCanti4.Text
                    xLote(WLugar, 9) = WLote5.Text
                    xLote(WLugar, 10) = WCanti5.Text
                    CargaLote.Visible = False
                    DBGrid1.Col = 5
                    DBGrid1.Text = "S"
                    If DBGrid1.Row < 40 Then
                       DBGrid1.Row = DBGrid1.Row + 1
                       WRow = DBGrid1.Row
                       XRow = DBGrid1.Row
                       DBGrid1.Col = 3
                       KeyCode = 0
                    End If
                    DBGrid1.Row = XRow
                    DBGrid1.Col = 3
                    KeyCode = 0
                    DBGrid1.SetFocus
                    Exit Sub
                        Else
                    WLote5.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti5.SetFocus
                    Else
                m$ = WTerminado.Text + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
            End If
            
            If WEstado = "N" Then
                WLote5.SetFocus
            End If
            
        End If
    
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo5 >= Val(WCanti5.Text) Then
            WCanti5.Text = Pusing("###,###.##", WCanti5.Text)
            Call Verifica_Lote
            If WEstado = "S" Then
                WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                xLote(WLugar, 1) = WLote1.Text
                xLote(WLugar, 2) = WCanti1.Text
                xLote(WLugar, 3) = WLote2.Text
                xLote(WLugar, 4) = WCanti2.Text
                xLote(WLugar, 5) = WLote3.Text
                xLote(WLugar, 6) = WCanti3.Text
                xLote(WLugar, 7) = WLote4.Text
                xLote(WLugar, 8) = WCanti4.Text
                xLote(WLugar, 9) = WLote5.Text
                xLote(WLugar, 10) = WCanti5.Text
                CargaLote.Visible = False
                DBGrid1.Col = 5
                DBGrid1.Text = "S"
                If DBGrid1.Row < 40 Then
                    DBGrid1.Row = DBGrid1.Row + 1
                    WRow = DBGrid1.Row
                    XRow = DBGrid1.Row
                    DBGrid1.Col = 3
                    KeyCode = 0
                End If
                DBGrid1.Row = XRow
                DBGrid1.Col = 3
                KeyCode = 0
                Exit Sub
            End If
                Else
            XSaldo5 = WSaldo5
            XSaldo5 = Pusing("###,###.##", XSaldo5)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo5
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            WLote5.SetFocus
        End If
        
        Rem WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
        Rem Call Verifica_Lote
        Rem If WEstado = "S" Then
        Rem     XLote(WRow, 1) = WLote1.Text
        Rem     XLote(WRow, 2) = WCanti1.Text
        Rem     XLote(WRow, 3) = WLote2.Text
        Rem     XLote(WRow, 4) = WCanti2.Text
        Rem     XLote(WRow, 5) = Wlote3.Text
        Rem     XLote(WRow, 6) = WCanti3.Text
        Rem     CargaLote.Visible = False
        Rem     DBGrid1.Col = 5
        Rem     DBGrid1.Text = "S"
        Rem     If DBGrid1.Row < 40 Then
        Rem         DBGrid1.Row = DBGrid1.Row + 1
        Rem         WRow = DBGrid1.Row
        Rem         XRow = DBGrid1.Row
        Rem         DBGrid1.Col = 3
        Rem         KeyCode = 0
        Rem     End If
        Rem     DBGrid1.Row = XRow
        Rem     DBGrid1.Col = 3
        Rem     KeyCode = 0
        Rem     Exit Sub
        Rem End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
ReDim UserData(0 To 5, 0 To 80)

mTotalRows& = 80

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
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Producto"
             DBGrid1.Columns(newcnt).Width = 1800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 3800
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Cantidad S/Pedido"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Cantidad a Entregar"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad a Restar"
             DBGrid1.Columns(newcnt).Width = 1600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "OK"
             DBGrid1.Columns(newcnt).Width = 300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = False
             DBGrid1.Columns(newcnt).Alignment = 1
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i

DBGrid1.Font.Bold = True

 WEnvase(1) = 20
 WEnvase(2) = 21
 WEnvase(3) = 22
 WEnvase(4) = 23
 WEnvase(5) = 24
 WEnvase(6) = 25
 WEnvase(7) = 26
 WEnvase(8) = 30
 WEnvase(9) = 28
 
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    For Cicla = 1 To 9
        spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
            rstEnvase.Close
                Else
            WImpre(Cicla) = ""
        End If
    Next Cicla
    
    Call Conecta_Empresa

    Erase XEnvase
    Erase xLote
    
    WCanti1.Text = ""
    WLote1.Text = ""
    WCanti2.Text = ""
    WLote2.Text = ""
    WCanti3.Text = ""
    WLote3.Text = ""
    WCanti4.Text = ""
    WLote4.Text = ""
    WCanti5.Text = ""
    WLote5.Text = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = 0
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pedido.SetFocus
     
End Sub

Private Sub Proceso_Click()

    Erase XEnvase
        
    For A = 0 To 7
    Suma = A * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next A
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    Erase ClavePedido
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    canti = !Cantidad - !Facturado
                    
                    If canti > 0 Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = !Terminado
                        Auxi1 = !Terminado
                
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                
                        Cantidad = IIf(IsNull(rstPedido!Cantidad1), "0", rstPedido!Cantidad1)
                        DBGrid1.Col = 3
                        DBGrid1.Text = Pusing("###,###.##", Str$(Cantidad))
                
                        Resta = IIf(IsNull(rstPedido!Cantidad2), "0", rstPedido!Cantidad2)
                        DBGrid1.Col = 4
                        DBGrid1.Text = Pusing("###,###.##", Str$(Resta))
                    
                        If Resta <> 0 Or Cantidad <> 0 Then
                            DBGrid1.Col = 5
                            DBGrid1.Text = "S"
                        End If
                    
                        WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                        
                        xLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        xLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        xLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        xLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        xLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        xLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        xLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        xLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        xLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        xLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                    
                        Envase1.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                        Canti1.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                        Envase2.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env2)
                        Canti2.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv2)
                        Envase3.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env3)
                        Canti3.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv3)
                        Envase4.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env4)
                        Canti4.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv4)
                        Envase5.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env5)
                        Canti5.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv5)
                    
                        Auxiliar(Renglon, 1) = Auxi1
                        Auxiliar(Renglon, 2) = canti
                        
                        ClavePedido(Renglon) = rstPedido!Clave
                    
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
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
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        descri2.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri3.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri4.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
                    
    spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnvases.RecordCount > 0 Then
        Descri5.Caption = rstEnvases!Abreviatura
        rstEnvases.Close
    End If
    
    Call Conecta_Empresa
    
    If Val(Envase1.Text) = 0 Then
        Envase1.Text = ""
    End If
    If Val(Envase2.Text) = 0 Then
        Envase2.Text = ""
    End If
    If Val(Envase3.Text) = 0 Then
        Envase3.Text = ""
    End If
    If Val(Envase4.Text) = 0 Then
        Envase4.Text = ""
    End If
    If Val(Envase5.Text) = 0 Then
        Envase5.Text = ""
    End If
    
    If Val(Canti1.Text) = 0 Then
        Canti1.Text = ""
    End If
    If Val(Canti2.Text) = 0 Then
        Canti2.Text = ""
    End If
    If Val(Canti3.Text) = 0 Then
        Canti3.Text = ""
    End If
    If Val(Canti4.Text) = 0 Then
        Canti4.Text = ""
    End If
    If Val(Canti5.Text) = 0 Then
        Canti5.Text = ""
    End If
    
    WRenglon = Renglon
    Renglon = 0
    
    For Da = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(Da, 1)
        canti = Auxiliar(Da, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) <> "PT" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(Auxi1, 3) + Right$(Auxi1, 7)
                ClavePreciosMp = Cliente.Text + Auxi1
                
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
                
                spPreciosMp = "ConsultaPreciosMp " + "'" + ClavePreciosMp + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
                If rstPreciosMp.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    Precio = rstPreciosMp!Precio
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
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 9 Step 2
                    If Val(xLote(Da, Ciclo)) = 0 Then
                        xLote(Da, Ciclo) = ""
                            Else
                        ZEntra = "N"
                        
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                        
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
                        
                        Call Conecta_Empresa
                            
                        Rem XParam = "'" + xLote(Da, Ciclo) + "'"
                        Rem spLaudo = "ListaLaudo " + XParam
                        Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstLaudo.RecordCount > 0 Then
                        Rem     xLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                        Rem     rstLaudo.Close
                        Rem End If
                        
                    End If
                Next Ciclo

                If Val(canti) <> 0 Then
                    WNeto = WNeto + (Val(canti) * Precio)
                End If
            
            Case Else
            
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
            
                spPrecios = "ConsultaPrecios " + "'" + ClavePrecios + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    Lugar1 = Int((Renglon - 1) / 10) * 10
                    Lugar2 = Renglon - Lugar1
                
                    DBGrid1.FirstRow = Lugar1
                    DBGrid1.Row = Lugar2 - 1
            
                    DBGrid1.Col = 1
                    DBGrid1.Text = rstPrecios!Descripcion
                
                    Rem DBGrid1.Col = 3
                    Rem DBGrid1.Text = Pusing("###,###.##", Str$(rstPrecios!Precio))
            
                    Precio = rstPrecios!Precio
                    rstPrecios.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 9 Step 2
                    If Val(xLote(Da, Ciclo)) = 0 Then
                        xLote(Da, Ciclo) = ""
                    End If
                Next Ciclo

                If Val(canti) <> 0 Then
                    WNeto = WNeto + (Val(canti) * Precio)
                End If
                
        End Select
        
    Next Da
    
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

End Sub

Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spPedido = "ConsultaPedido1 " + "'" + Pedido.Text + "'"
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            If rstPedido!Autorizo <> "X" Then
                rstPedido.Close
                Call Conecta_Empresa
             
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                A% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                
                    Else
                    
                WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                Cliente.Text = rstPedido!Cliente
                Fecha.Text = rstPedido!Fecha
                WFecEntrega = rstPedido!FecEntrega
                WObservaciones = rstPedido!Observaciones
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
                
                rstPedido.Close
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    Cliente.Text = rstCliente!Cliente
                    DesCliente.Caption = rstCliente!Razon
                    
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
                
                Call Conecta_Empresa
                
                Call Proceso_Click
                DBGrid1.FirstRow = 0
                DBGrid1.Row = 0
                DBGrid1.Col = 3
                DBGrid1.SetFocus
            End If
            
                Else
            
            Call Conecta_Empresa
            
        End If
        
    End If
End Sub

Private Sub Envase1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + Envase1.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri1.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti1.SetFocus
                Else
            Envase1.SetFocus
        End If
        
        Call Conecta_Empresa
        
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
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + Envase2.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            descri2.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti2.SetFocus
                Else
            Envase2.SetFocus
        End If
        
        Call Conecta_Empresa
        
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
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + Envase3.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri3.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti3.SetFocus
                Else
            Envase3.SetFocus
        End If
        
        Call Conecta_Empresa
        
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
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + Envase4.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri4.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti4.SetFocus
                Else
            Envase4.SetFocus
        End If
        
        Call Conecta_Empresa
        
        
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
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnvases = "ConsultaEnvases " + "'" + Envase5.Text + "'"
        Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvases.RecordCount > 0 Then
            Descri5.Caption = rstEnvases!Abreviatura
            rstEnvases.Close
            Canti5.SetFocus
                Else
            Envase5.SetFocus
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Envase1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
    If WLote3.Text <> "" Then
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
        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
    End If
    
    If WEstado = "S" Then
    
        Erase ControlLote
        ControlLote(1, 1) = WLote1.Text
        ControlLote(1, 2) = WCanti1.Text
        ControlLote(2, 1) = WLote2.Text
        ControlLote(2, 2) = WCanti2.Text
        ControlLote(3, 1) = WLote3.Text
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
                            A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
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
        ControlLote(3, 1) = WLote3.Text
        ControlLote(3, 2) = WCanti3.Text
        ControlLote(4, 1) = WLote4.Text
        ControlLote(4, 2) = WCanti4.Text
        ControlLote(5, 1) = WLote5.Text
        ControlLote(5, 2) = WCanti5.Text
    
        For Ciclo1 = 1 To 5
    
            WLote = ControlLote(Ciclo1, 1)
            WCanti = Val(ControlLote(Ciclo1, 2))
            
            If WLote <> "" Or Val(WCanti) <> 0 Then
            
            If Left$(XTerminado, 2) <> "PT" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                        Select Case WTipoPedido
                            Case "PG", "CO"
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case "FA"
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            Case Else
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        End Select
                    End If
                    
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
                                A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
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
                                    A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
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
                    
                    Call Conecta_Empresa
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                    
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
                    
                    Call Conecta_Empresa
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
                        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Then
                            Select Case WTipoPedido
                                Case "PG", "CO"
                                    WEmpresa = "0001"
                                    txtOdbc = "Empresa01"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case "FA"
                                    WEmpresa = "0005"
                                    txtOdbc = "Empresa05"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                Case Else
                                    WEmpresa = "0007"
                                    txtOdbc = "Empresa07"
                                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            End Select
                        End If
                    
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
                                A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
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
                            Rem WVida = 0
                            Rem
                            Rem spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                            Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                            Rem If rstTerminado.RecordCount > 0 Then
                            Rem     WVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                            Rem     rstTerminado.Close
                            Rem End If
                            Rem
                            Rem If WVida <> 0 Then
                            Rem
                            Rem     WMes = Val(Mid$(WFechaHoja, 4, 2))
                            Rem     WAno = Val(Right$(WFechaHoja, 4))
                            Rem     For Ciclo = 1 To WVida
                            Rem         WMes = WMes + 1
                            Rem         If WMes > 12 Then
                            Rem             WAno = WAno + 1
                            Rem             WMes = 1
                            Rem         End If
                            Rem     Next Ciclo
                            Rem     XMes = Str$(WMes)
                            Rem     XAno = Str$(WAno)
                            Rem     Call Ceros(XMes, 2)
                            Rem     Call Ceros(XAno, 4)
                            Rem     WVencimiento = "01/" + XMes + "/" + XAno
                            Rem
                            Rem     WFechaActual = "01" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            Rem     WFechaActualOrd = Right$(WFechaActual, 4) + Mid$(WFechaActual, 4, 2) + Left$(WFechaActual, 2)
                            Rem
                            Rem     WFechaVencimiento = "01" + Mid$(WVencimiento, 3, 10)
                            Rem     WFechaVencimientoOrd = Right$(WFechaVencimiento, 4) + Mid$(WFechaVencimiento, 4, 2) + Left$(WFechaVencimiento, 2)
                            Rem
                            Rem     Pasa = "S"
                            Rem     If WFechaActualOrd >= WFechaVencimientoOrd Then
                            Rem         Pasa = "N"
                            Rem             Else
                            Rem         Meses = 0
                            Rem         WMes = Val(Mid$(WFechaActual, 4, 2))
                            Rem         WAno = Val(Right$(WFechaActual, 4))
                            Rem         Do
                            Rem             Meses = Meses + 1
                            Rem             WMes = WMes + 1
                            Rem             If WMes > 12 Then
                            Rem                 WAno = WAno + 1
                            Rem                 WMes = 1
                            Rem             End If
                            Rem             XMes = Str$(WMes)
                            Rem             XAno = Str$(WAno)
                            Rem             Call Ceros(XMes, 2)
                            Rem             Call Ceros(XAno, 4)
                            Rem             WCompara = "01/" + XMes + "/" + XAno
                            Rem             If WCompara = WFechaVencimiento Then
                            Rem                 Exit Do
                            Rem             End If
                            Rem         Loop
                            Rem         If Meses <= 12 Then
                            Rem             Pasa = "N"
                            Rem         End If
                            Rem     End If
                            Rem
                            Rem     If Pasa = "N" Then
                            Rem         m$ = "EL Producto tiene menos de un año de vida util"
                            Rem         G% = MsgBox(m$, 0, "Actualizacion de Pedido")
                            Rem         WEstado = "N"
                            Rem     End If
                            Rem
                            Rem End If
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
                                    A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
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
                                
                        Call Conecta_Empresa
                        
                
                            Else
                            
                        WEntra = "S"
                        
                    End If
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        A = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
            End Select
            
            End If
            
        Next Ciclo1

    End If
    
End Sub

Private Sub reImpre_Click()

    With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                WAuxiliar = !Nombre
            End If
    End With
    
    If Val(WEmpresa) = 1 Then
        Rem Open "dada.txt" For Output As #1
        Open "lpt2" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
            Else
        Rem Open "dada.txt" For Output As #1
        Open "lpt1" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "11" + Chr$(72);
    End If

        Print #1, Tab(1); String$(79, "-")
        
        Print #1, Tab(1); "|";
        Print #1, Tab(5); WAuxiliar;
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Pedido.Text;
        Print #1, " / ";
        Print #1, WVersion;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Cliente";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Cliente.Text;
        Print #1, Tab(40); Left$(DesCliente.Caption, 35);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Fecha.Text;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Ent.";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WFecEntrega;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "C.Pago";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDespago;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Entrega";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDirentrega;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WObservaciones;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|";
        Print #1, Tab(2); "Producto";
        Print #1, Tab(16); "|";
        Print #1, Tab(17); "Descripcion";
        Print #1, Tab(47); "|";
        Print #1, Tab(48); "Partida";
        Print #1, Tab(58); "|";
        Print #1, Tab(59); "Cantidad";
        Print #1, Tab(67); "|";
        Print #1, Tab(68); "Envase";
        Print #1, Tab(80); "|"
        Print #1, Tab(1); String$(79, "-")
        
        XLinea = 0
        WCounter = 0
                    
        For A = 0 To 7
        
                Suma = A * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WCounter = WCounter + 1
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    
                    If DBGrid1.Text <> "" Then
                    
                    WArticulo = DBGrid1.Text
                    
                    DBGrid1.Col = 1
                    WDescripcion = DBGrid1.Text
                    
                    DBGrid1.Col = 2
                    WCantidad = Val(DBGrid1.Text)
                    
                    If WCantidad <> 0 Then
                    
                        Print #1, Tab(1); "|";
                        Print #1, Tab(2); WArticulo;
                        Print #1, Tab(16); "|";
                        Print #1, Tab(17); Left$(WDescripcion, 28);
                        Print #1, Tab(47); "|";
                        Print #1, Tab(58); "|";
                        Print #1, Tab(59); Alinea("#####.##", Str$(WCantidad));
                        Print #1, Tab(67); "|";
                                
                        For Cicla = 1 To 6 Step 2
                            If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                                Select Case Cicla
                                    Case 1
                                        XEmpresa = WEmpresa
                                        Select Case Val(WEmpresa)
                                            Case 1, 3, 5, 6, 7, 10
                                                WEmpresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                                WEmpresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        End Select
                                    
                                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvase.RecordCount > 0 Then
                                            WAbre = rstEnvase!Abreviatura
                                            rstEnvase.Close
                                                Else
                                            WAbre = ""
                                        End If
                                        
                                        Call Conecta_Empresa
                                            
                                        Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                        Print #1, Tab(80); "|"

                                    Case Else
                                        Print #1, Tab(1); "|";
                                        Print #1, Tab(16); "|";
                                        Print #1, Tab(47); "|";
                                        Print #1, Tab(58); "|";
                                        Print #1, Tab(67); "|";
                                                
                                        XEmpresa = WEmpresa
                                        Select Case Val(WEmpresa)
                                            Case 1, 3, 5, 6, 7, 10
                                                WEmpresa = "0001"
                                                txtOdbc = "Empresa01"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                            Case Else
                                                WEmpresa = "0008"
                                                txtOdbc = "Empresa08"
                                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                        End Select
                                                
                                        spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                        If rstEnvase.RecordCount > 0 Then
                                            WAbre = rstEnvase!Abreviatura
                                            rstEnvase.Close
                                                Else
                                            WAbre = ""
                                        End If
                                        
                                        Call Conecta_Empresa
                                            
                                        Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                        Print #1, Tab(80); "|"
                                        XLinea = XLinea + 1
                                            
                                End Select
                            End If
                        Next Cicla
                        XLinea = XLinea + 1

                    End If
                    
                    End If
                                        
                Next iRow
            
        Next A
        
        For WDa = XLinea To 10
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(47); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|"; WImpre(1);
        Print #1, Tab(10); "|"; WImpre(2);
        Print #1, Tab(18); "|"; WImpre(3);
        Print #1, Tab(26); "|"; WImpre(4);
        Print #1, Tab(34); "|"; WImpre(5);
        Print #1, Tab(42); "|"; WImpre(6);
        Print #1, Tab(50); "|"; WImpre(7);
        Print #1, Tab(58); "|"; WImpre(8);
        Print #1, Tab(66); "|"; WImpre(9);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|(020)";
        Print #1, Tab(10); "|(021)";
        Print #1, Tab(18); "|(022)";
        Print #1, Tab(26); "|(023)";
        Print #1, Tab(34); "|(024)";
        Print #1, Tab(42); "|(025)";
        Print #1, Tab(50); "|(026)";
        Print #1, Tab(58); "|(030)";
        Print #1, Tab(66); "|(028)";
        Print #1, Tab(80); "|"
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""
    
    Print #1, Chr$(12)
    Close #1
    
    DBGrid1.FirstRow = 0
    DBGrid1.Row = 0
    DBGrid1.Col = 0

End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            spEnvases = "ListaEnvases"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstEnvases
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnvases!Envases) + " " + rstEnvases!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstEnvases!Envases
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnvases.Close
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            Indice = Pantalla.ListIndex
            WEnvases = WIndice.List(Indice)
            spEnvases = "ConsultaEnvases " + "'" + Str$(WEnvases) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
            
                Entra = "N"
                
                If Val(Envase1.Text) = 0 And Entra = "N" Then
                    Envase1.Text = rstEnvases!Envases
                    Descri1.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti1.SetFocus
                End If
                
                If Val(Envase2.Text) = 0 And Entra = "N" Then
                    Envase2.Text = rstEnvases!Envases
                    descri2.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti2.SetFocus
                End If
                
                If Val(Envase3.Text) = 0 And Entra = "N" Then
                    Envase3.Text = rstEnvases!Envases
                    Descri3.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti3.SetFocus
                End If
                
                If Val(Envase4.Text) = 0 And Entra = "N" Then
                    Envase4.Text = rstEnvases!Envases
                    Descri4.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti4.SetFocus
                End If
                
                If Val(Envase5.Text) = 0 And Entra = "N" Then
                    Envase5.Text = rstEnvases!Envases
                    Descri5.Caption = rstEnvases!Descripcion
                    Entra = "S"
                    Canti5.SetFocus
                End If
                
                rstEnvases.Close
                    
            End If
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
    
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




