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
   Begin VB.ComboBox MarcaFactura 
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
      Left            =   1320
      TabIndex        =   70
      Top             =   840
      Width           =   2055
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
      Height          =   2655
      Left            =   1200
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   7575
      Begin VB.TextBox WBultos1 
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
         Left            =   6360
         TabIndex        =   67
         Text            =   " "
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox WBultos2 
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
         Left            =   6360
         TabIndex        =   66
         Text            =   " "
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox WBultos3 
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
         Left            =   6360
         TabIndex        =   65
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox WBultos4 
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
         Left            =   6360
         TabIndex        =   64
         Text            =   " "
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox WBultos5 
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
         Left            =   6360
         TabIndex        =   63
         Text            =   " "
         Top             =   2040
         Width           =   855
      End
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bultos"
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
         Left            =   6360
         TabIndex        =   68
         Top             =   240
         Width           =   855
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      ItemData        =   "modped.frx":0000
      Left            =   3360
      List            =   "modped.frx":0007
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   8055
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "modped.frx":0015
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
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   120
      TabIndex        =   69
      Top             =   840
      Width           =   1095
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
      Visible         =   0   'False
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
Private WMarcaFactura As Integer

Private WImpre(10) As String
Private WEnvase(10) As String
Private Envase(5, 2) As String
Private Auxiliar(100, 14) As String
Private ClavePedido(100) As String
Private BajaLote(5, 2) As String
Private XLote(100, 30) As String
Private XEnvase(100, 6) As String
Private ImpreEnvase(10) As String
Private EmiteCerti(1000, 3) As String
Private CargaEmpresa(12, 2) As String
Private LugarCerti As Integer
Private TipoEnvase(100, 2) As String

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
Dim XBultos1 As String
Dim XEnv2 As String
Dim XCantiEnv2 As String
Dim XBultos2 As String
Dim XEnv3 As String
Dim XCantiEnv3 As String
Dim XBultos3 As String
Dim XEnv4 As String
Dim XCantiEnv4 As String
Dim XBultos4 As String
Dim XEnv5 As String
Dim XCantiEnv5 As String
Dim XBultos5 As String
Dim XMes As String
Dim XAno As String

Dim ControlLote(5, 2) As String
Dim WSaldo As Double
Dim WCanti As Double
Dim WLote As String
Dim WLugar As Integer

Dim ZZGrilla(100, 15) As String
Dim ZZHoja(100) As String
Dim ZZNumeroHoja As String

Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim XEspecificaciones(100) As String
Dim ZVector(100, 11) As String
Dim WEspecif(100) As String

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
    
    XLote(WLugar, 1) = ""
    XLote(WLugar, 2) = ""
    XLote(WLugar, 3) = ""
    XLote(WLugar, 4) = ""
    XLote(WLugar, 5) = ""
    XLote(WLugar, 6) = ""
    XLote(WLugar, 7) = ""
    XLote(WLugar, 8) = ""
    XLote(WLugar, 9) = ""
    XLote(WLugar, 10) = ""
    XLote(WLugar, 11) = ""
    XLote(WLugar, 12) = ""
    XLote(WLugar, 13) = ""
    XLote(WLugar, 14) = ""
    XLote(WLugar, 15) = ""
    XLote(WLugar, 16) = ""
    XLote(WLugar, 17) = ""
    XLote(WLugar, 18) = ""
    XLote(WLugar, 19) = ""
    XLote(WLugar, 20) = ""
    XLote(WLugar, 21) = ""
    XLote(WLugar, 22) = ""
    XLote(WLugar, 23) = ""
    XLote(WLugar, 24) = ""
    XLote(WLugar, 25) = ""
    
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
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
    
    WRenglon = 0
    DBGrid1.Refresh
        
    For a = 0 To 9
    
        Suma = a * 10
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
                    a = MsgBox(m$, 0, "MODULO DE FACTURACION")
                    DBGrid1.Refresh
                    Exit Sub
                End If
            End If
               
        Next iRow
            
    Next a
    
    Erase TipoEnvase
    LugarEnvase = 0
    SumaEnvases = 0
    
    For CicloEnvase = 1 To 100
    
        SumaEnvases = SumaEnvases + Val(XLote(CicloEnvase, 12))
        SumaEnvases = SumaEnvases + Val(XLote(CicloEnvase, 14))
        SumaEnvases = SumaEnvases + Val(XLote(CicloEnvase, 16))
        SumaEnvases = SumaEnvases + Val(XLote(CicloEnvase, 18))
        SumaEnvases = SumaEnvases + Val(XLote(CicloEnvase, 20))
        
        Rem envase posicion 1
        Entra = "S"
        For CicloEnvaseII = 1 To LugarEnvase
            If TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 11) Then
                TipoEnvase(CicloEnvaseII, 2) = Str$(Val(TipoEnvase(CicloEnvaseII, 2)) + Val(XLote(CicloEnvase, 12)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            LugarEnvase = LugarEnvase + 1
            TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 11)
            TipoEnvase(CicloEnvaseII, 2) = XLote(CicloEnvase, 12)
        End If
        
        Rem envase posicion 2
        Entra = "S"
        For CicloEnvaseII = 1 To LugarEnvase
            If TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 13) Then
                TipoEnvase(CicloEnvaseII, 2) = Str$(Val(TipoEnvase(CicloEnvaseII, 2)) + Val(XLote(CicloEnvase, 14)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            LugarEnvase = LugarEnvase + 1
            TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 13)
            TipoEnvase(CicloEnvaseII, 2) = XLote(CicloEnvase, 14)
        End If
        
        Rem envase posicion 3
        Entra = "S"
        For CicloEnvaseII = 1 To LugarEnvase
            If TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 15) Then
                TipoEnvase(CicloEnvaseII, 2) = Str$(Val(TipoEnvase(CicloEnvaseII, 2)) + Val(XLote(CicloEnvase, 16)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            LugarEnvase = LugarEnvase + 1
            TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 15)
            TipoEnvase(CicloEnvaseII, 2) = XLote(CicloEnvase, 16)
        End If
        
        Rem envase posicion 4
        Entra = "S"
        For CicloEnvaseII = 1 To LugarEnvase
            If TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 17) Then
                TipoEnvase(CicloEnvaseII, 2) = Str$(Val(TipoEnvase(CicloEnvaseII, 2)) + Val(XLote(CicloEnvase, 18)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            LugarEnvase = LugarEnvase + 1
            TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 17)
            TipoEnvase(CicloEnvaseII, 2) = XLote(CicloEnvase, 18)
        End If
        
        Rem envase posicion 5
        Entra = "S"
        For CicloEnvaseII = 1 To LugarEnvase
            If TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 19) Then
                TipoEnvase(CicloEnvaseII, 2) = Str$(Val(TipoEnvase(CicloEnvaseII, 2)) + Val(XLote(CicloEnvase, 20)))
                Entra = "N"
                Exit For
            End If
        Next CicloEnvaseII
        
        If Entra = "S" Then
            LugarEnvase = LugarEnvase + 1
            TipoEnvase(CicloEnvaseII, 1) = XLote(CicloEnvase, 19)
            TipoEnvase(CicloEnvaseII, 2) = XLote(CicloEnvase, 20)
        End If
        
    Next CicloEnvase
    
    Rem SumaEnvases = Val(Canti1.Text) + Val(Canti2.Text) + Val(Canti3.Text) + Val(Canti4.Text) + Val(Canti5.Text)
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
        
        If Val(TipoEnvase(1, 2)) <> 0 Then
            ZDescri1 = ""
            spEnvases = "ConsultaEnvases " + "'" + TipoEnvase(1, 1) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri1 = rstEnvases!Abreviatura
                rstEnvases.Close
            End If
            m$ = m$ + Str$(Val(TipoEnvase(1, 2))) + " envases de " + ZDescri1 + Chr$(13) + ""
        End If
        
        If Val(TipoEnvase(2, 2)) <> 0 Then
            ZDescri1 = ""
            spEnvases = "ConsultaEnvases " + "'" + TipoEnvase(2, 1) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri2 = rstEnvases!Abreviatura
                rstEnvases.Close
            End If
            m$ = m$ + Str$(Val(TipoEnvase(2, 2))) + " envases de " + ZDescri2 + Chr$(13) + ""
        End If
        
        If Val(TipoEnvase(3, 2)) <> 0 Then
            ZDescri1 = ""
            spEnvases = "ConsultaEnvases " + "'" + TipoEnvase(3, 1) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri3 = rstEnvases!Abreviatura
                rstEnvases.Close
            End If
            m$ = m$ + Str$(Val(TipoEnvase(3, 2))) + " envases de " + ZDescri3 + Chr$(13) + ""
        End If
        
        If Val(TipoEnvase(4, 2)) <> 0 Then
            ZDescri1 = ""
            spEnvases = "ConsultaEnvases " + "'" + TipoEnvase(4, 1) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri4 = rstEnvases!Abreviatura
                rstEnvases.Close
            End If
            m$ = m$ + Str$(Val(TipoEnvase(4, 2))) + " envases de " + ZDescri4 + Chr$(13) + ""
        End If
        
        If Val(TipoEnvase(5, 2)) <> 0 Then
            ZDescri1 = ""
            spEnvases = "ConsultaEnvases " + "'" + TipoEnvase(5, 1) + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                ZDescri5 = rstEnvases!Abreviatura
                rstEnvases.Close
            End If
            m$ = m$ + Str$(Val(TipoEnvase(5, 2))) + " envases de " + ZDescri5 + Chr$(13) + ""
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
    LugarHoja = 0
    Erase ZZHoja
        
    For a = 0 To 9
        
        Suma = a * 10
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
                    
                XLote1 = XLote(WLugar, 1)
                XLote2 = XLote(WLugar, 3)
                XLote3 = XLote(WLugar, 5)
                XLote4 = XLote(WLugar, 7)
                XLote5 = XLote(WLugar, 9)
                    
                If Left$(Articulo, 2) <> "PT" And Left$(Articulo, 2) <> "PE" And Left$(Articulo, 2) <> "YQ" And Left$(Articulo, 2) <> "YF" And Left$(Articulo, 2) <> "YP" And Left$(Articulo, 2) <> "YH" Then
                
                    For Ciclo = 1 To 9 Step 2
                    
                        If XLote(WLugar, Ciclo) <> "" Then
                            
                            ZEntra = "N"
                            
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
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Laudo"
                            ZSql = ZSql + " Where Laudo.PartiOri = " + "'" + XLote(WLugar, Ciclo) + "'"
                            ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            spLaudo = ZSql
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
                            
                                ZZCodigo = Left$(Articulo, 3) + Mid$(Articulo, 6, 10)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Guia"
                                ZSql = ZSql + " Where Guia.PartiOri = " + "'" + XLote(WLugar, Ciclo) + "'"
                                ZSql = ZSql + " and Guia.Articulo = " + "'" + ZZCodigo + "'"
                                ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                spMovguia = ZSql
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
                                            
                    XLote1 = XLote(WLugar, 1)
                    XLote2 = XLote(WLugar, 3)
                    XLote3 = XLote(WLugar, 5)
                    XLote4 = XLote(WLugar, 7)
                    XLote5 = XLote(WLugar, 9)
                
                End If
                    
                XCantiLote1 = XLote(WLugar, 2)
                XCantiLote2 = XLote(WLugar, 4)
                XCantiLote3 = XLote(WLugar, 6)
                XCantiLote4 = XLote(WLugar, 8)
                XCantiLote5 = XLote(WLugar, 10)
                
                XEnv1 = XLote(WLugar, 11)
                XCantiEnv1 = XLote(WLugar, 12)
                XBultos1 = XLote(WLugar, 21)
                XEnv2 = XLote(WLugar, 13)
                XCantiEnv2 = XLote(WLugar, 14)
                XBultos2 = XLote(WLugar, 22)
                XEnv3 = XLote(WLugar, 15)
                XCantiEnv3 = XLote(WLugar, 16)
                XBultos2 = XLote(WLugar, 23)
                XEnv4 = XLote(WLugar, 17)
                XCantiEnv4 = XLote(WLugar, 18)
                XBultos3 = XLote(WLugar, 24)
                XEnv5 = XLote(WLugar, 19)
                XCantiEnv5 = XLote(WLugar, 20)
                XBultos4 = XLote(WLugar, 25)
                
                Rem XEnv1 = Envase1.Text
                Rem XCantiEnv1 = Canti1.Text
                Rem XEnv2 = Envase2.Text
                Rem XCantiEnv2 = Canti2.Text
                Rem XEnv3 = Envase3.Text
                Rem XCantiEnv3 = Canti3.Text
                Rem XEnv4 = Envase4.Text
                Rem XCantiEnv4 = Canti4.Text
                Rem XEnv5 = Envase5.Text
                Rem XCantiEnv5 = Canti5.Text
                
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
                    Case 1, 3, 5, 6, 7, 10, 11
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
                ZSql = ZSql & "CantidadFac = " + "'" + "0" + "',"
                ZSql = ZSql & "Bultos1 = " + "'" + XBultos1 + "',"
                ZSql = ZSql & "Bultos2 = " + "'" + XBultos2 + "',"
                ZSql = ZSql & "Bultos3 = " + "'" + XBultos3 + "',"
                ZSql = ZSql & "Bultos4 = " + "'" + XBultos4 + "',"
                ZSql = ZSql & "Bultos5 = " + "'" + XBultos5 + "',"
                ZSql = ZSql & "FechaActualizacion = " + "'" + ZFechaActualiza + "',"
                ZSql = ZSql & "OrdFechaActualizacion = " + "'" + ZOrdFechaActualiza + "'"
                ZSql = ZSql & " Where Clave = " + "'" + WClavePedido + "'"
                
                spPedido = ZSql
                
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql & "UPDATE Pedido SET "
                ZSql = ZSql & " MarcaFactura = " + "'" + Trim(Str$(MarcaFactura.ListIndex)) + "'"
                ZSql = ZSql & " Where Pedido = " + "'" + Pedido.Text + "'"
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
                
                If Val(XLote1) <> 0 Then
                    LugarHoja = LugarHoja + 1
                    ZZHoja(LugarHoja) = XLote1
                End If
                
                If Val(XLote2) <> 0 Then
                    LugarHoja = LugarHoja + 1
                    ZZHoja(LugarHoja) = XLote2
                End If
                
                If Val(XLote3) <> 0 Then
                    LugarHoja = LugarHoja + 1
                    ZZHoja(LugarHoja) = XLote3
                End If
                
                If Val(XLote4) <> 0 Then
                    LugarHoja = LugarHoja + 1
                    ZZHoja(LugarHoja) = XLote4
                End If
                    
                If Val(XLote5) <> 0 Then
                    LugarHoja = LugarHoja + 1
                    ZZHoja(LugarHoja) = XLote5
                End If
                
            End If
                                        
        Next iRow
            
    Next a
    
    If Cliente.Text = "T00140" Then
    
        T$ = "HOJAS DE PRODUCCION"
        m$ = "Desea enviar por email las Hojas de Produccion a TANATEX ?"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            Sql1 = "DELETE ImpreHoja44000"
            spImpreHoja = Sql1
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
        
            For CicloHoja = 1 To 100
                If Val(ZZHoja(CicloHoja)) <> 0 Then
                    ZZNumeroHoja = ZZHoja(CicloHoja)
                    Call Envio_Hoja_Email
                End If
            Next CicloHoja
    
            Listado.WindowTitle = "Impresion de Hoja de Produccion"
            Listado.WindowTop = 0
            Listado.WindowLeft = 0
            Listado.WindowWidth = Screen.Width
            Listado.WindowHeight = Screen.Height
        
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.ReportFileName = "HojaProduccion.rpt"
            Listado.GroupSelectionFormula = "{ImpreHoja44000.Hoja} in 0 to 999999"
        
            Listado.SQLQuery = "SELECT ImpreHoja44000.Hoja, ImpreHoja44000.Renglon, ImpreHoja44000.Fecha, ImpreHoja44000.Codigo1, ImpreHoja44000.Codigo2, ImpreHoja44000.Articulo1, ImpreHoja44000.Articulo2, ImpreHoja44000.Descripcion, ImpreHoja44000.Canti1, ImpreHoja44000.Lote1, ImpreHoja44000.Canti2, ImpreHoja44000.Lote2, ImpreHoja44000.Canti3, ImpreHoja44000.Lote3, ImpreHoja44000.Teorico, ImpreHoja44000.CantidadReal, ImpreHoja44000.VersionI, ImpreHoja44000.VersionII, ImpreHoja44000.VersionIII, ImpreHoja44000.LoteOri1, ImpreHoja44000.LoteOri2, ImpreHoja44000.LoteOri3, ImpreHoja44000.Nombre " _
                + "From " _
                + DSQ + ".dbo.ImpreHoja44000 ImpreHoja44000 " _
                + "WHERE " _
                + "ImpreHoja44000.Hoja >= 0 AND " _
                + "ImpreHoja44000.Hoja <= 999999"
                
            Listado.EMailToList = "CAROLINA.PENZ@TANATEXCHEMICALS.COM; drodriguez@surfactan.com.ar"
            Listado.EMailSubject = "Hojas de Produccion de SURFACTAN S.A."
            Listado.EMailMessage = "Les envio las hojas de produccion del material a remitir"
            
            Listado.Destination = 3
            Listado.PrintFileName = "Hojas.doc"
            Listado.PrintFileType = crptWinWord
            
            MiRuta = CurDir + "\"
            MiRutaII = Left$(CurDir, 1)
        
            Listado.Connect = Connect()
            Listado.Action = 1
            
            ChDrive MiRutaII
            ChDir MiRuta
            
        End If
    End If
        
    Call Limpia_Click

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Exit Sub

WError:
     Resume Next
        
End Sub

Sub Envio_Hoja_Email()
        
    WHoja = ZZNumeroHoja
    
    spHoja = "ListaHoja " + "'" + WHoja + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        WFecha = rstHoja!Fecha
        WCantidadReal = Str$(rstHoja!Real)
        WTeorico = Str$(rstHoja!Teorico)
        WProducto = rstHoja!Producto
        WVersionI = IIf(IsNull(rstHoja!VersionI), "", rstHoja!VersionI)
        WVersionII = IIf(IsNull(rstHoja!VersionII), "", rstHoja!VersionII)
        WVersionIII = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
        rstHoja.Close
    End If
    
    WNombre = ""
    spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WNombre = rstTerminado!Descripcion
        rstTerminado.Close
    End If
    
    
    WCodigo1 = Left$(WProducto, 2)
    WCodigo2 = Mid$(WProducto, 4, 5) + "/" + Right$(WProducto, 3)
    
    ZZRenglon = 0
    Erase ZZGrilla
    
    spHoja = "ListaHoja " + "'" + WHoja + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZZRenglon = ZZRenglon + 1
                
                    ZZGrilla(ZZRenglon, 1) = rstHoja!Tipo
                    ZZGrilla(ZZRenglon, 2) = rstHoja!Terminado
                    ZZGrilla(ZZRenglon, 3) = rstHoja!Articulo
                    ZZGrilla(ZZRenglon, 5) = Pusing("###,###.##", rstHoja!Cantidad)
                    ZZGrilla(ZZRenglon, 6) = Str$(rstHoja!Canti1)
                    ZZGrilla(ZZRenglon, 7) = Str$(rstHoja!lote1)
                    ZZGrilla(ZZRenglon, 8) = Str$(rstHoja!Canti2)
                    ZZGrilla(ZZRenglon, 9) = Str$(rstHoja!lote2)
                    ZZGrilla(ZZRenglon, 10) = Str$(rstHoja!Canti3)
                    ZZGrilla(ZZRenglon, 11) = Str$(rstHoja!lote3)
                    
                    ZSuma = rstHoja!Canti1 + rstHoja!Canti2 + rstHoja!Canti3
                    If ZSuma = 0 Then
                        ZZGrilla(ZZRenglon, 6) = Str$(rstHoja!Cantidad)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For DA = 1 To ZZRenglon
            
        Tipo = ZZGrilla(DA, 1)
        Auxi1 = ZZGrilla(DA, 2)
        Auxi2 = ZZGrilla(DA, 3)
                
        Select Case Tipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Auxi1 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZGrilla(DA, 4) = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Auxi2 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZGrilla(DA, 4) = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            Case Else
        End Select
    Next DA

    ZZLugar = 0

    For a = 1 To 40
    
        WTipo = UCase(ZZGrilla(a, 1))
        
        If Trim(WTipo) <> "" Then
        
            ZZLugar = ZZLugar + 1
            
            WTerminado = UCase(ZZGrilla(a, 2))
            WArticulo = UCase(ZZGrilla(a, 3))
            WDescripcion = UCase(ZZGrilla(a, 4))
            WCantidad = ZZGrilla(a, 5)
            WLinea = Str$(a)
            
            If WTipo = "M" Then
                WArticulo1 = Left$(WArticulo, 2)
                WArticulo2 = Mid$(WArticulo, 4, 3) + "-" + Right$(WArticulo, 3)
                    Else
                WArticulo1 = Left$(WTerminado, 2)
                WArticulo2 = Mid$(WTerminado, 4, 5) + "-" + Right$(WTerminado, 3)
            End If
            
            WCanti1 = ZZGrilla(a, 6)
            WLote1 = ZZGrilla(a, 7)
            WLoteOri1 = ""
            WCanti2 = ZZGrilla(a, 8)
            WLote2 = ZZGrilla(a, 9)
            WLoteOri2 = ""
            WCanti3 = ZZGrilla(a, 10)
            Wlote3 = ZZGrilla(a, 11)
            WLoteOri3 = ""
        
            For ZZPasalote = 1 To 3
                
                Select Case ZZPasalote
                    Case 1
                        XXLote = ZZGrilla(a, 7)
                    Case 2
                        XXLote = ZZGrilla(a, 9)
                    Case Else
                        XXLote = ZZGrilla(a, 11)
                End Select
         
                If WTipo = "M" Then
                    XParam = "'" + XXLote + "','" _
                                + WArticulo + "'"
                    spLaudo = "ListaLaudoArticulo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        Select Case ZZPasalote
                            Case 1
                                WLoteOri1 = rstLaudo!PartiOri
                            Case 2
                                WLoteOri2 = rstLaudo!PartiOri
                            Case Else
                                WLoteOri3 = rstLaudo!PartiOri
                        End Select
                        rstLaudo.Close
                    End If
                End If
                
            Next ZZPasalote
                
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ImpreHoja44000 ("
            ZSql = ZSql & "Hoja ,"
            ZSql = ZSql & "Renglon ,"
            ZSql = ZSql & "Fecha ,"
            ZSql = ZSql & "Codigo1 ,"
            ZSql = ZSql & "Codigo2 ,"
            ZSql = ZSql & "Nombre ,"
            ZSql = ZSql & "Articulo1 ,"
            ZSql = ZSql & "Articulo2 ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Canti1 ,"
            ZSql = ZSql & "Lote1 ,"
            ZSql = ZSql & "LoteOri1 ,"
            ZSql = ZSql & "Canti2 ,"
            ZSql = ZSql & "Lote2 ,"
            ZSql = ZSql & "LoteOri2 ,"
            ZSql = ZSql & "Canti3 ,"
            ZSql = ZSql & "Lote3 ,"
            ZSql = ZSql & "LoteOri3 ,"
            ZSql = ZSql & "Teorico ,"
            ZSql = ZSql & "CantidadReal ,"
            ZSql = ZSql & "VersionI ,"
            ZSql = ZSql & "VersionII ,"
            ZSql = ZSql & "VersionIII )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + WHoja + "',"
            ZSql = ZSql & "'" + WLinea + "',"
            ZSql = ZSql & "'" + WFecha + "',"
            ZSql = ZSql & "'" + WCodigo1 + "',"
            ZSql = ZSql & "'" + WCodigo2 + "',"
            ZSql = ZSql & "'" + WNombre + "',"
            ZSql = ZSql & "'" + WArticulo1 + "',"
            ZSql = ZSql & "'" + WArticulo2 + "',"
            ZSql = ZSql & "'" + WCantidad + "',"
            ZSql = ZSql & "'" + WDescripcion + "',"
            ZSql = ZSql & "'" + WCanti1 + "',"
            ZSql = ZSql & "'" + WLote1 + "',"
            ZSql = ZSql & "'" + WLoteOri1 + "',"
            ZSql = ZSql & "'" + WCanti2 + "',"
            ZSql = ZSql & "'" + WLote2 + "',"
            ZSql = ZSql & "'" + WLoteOri2 + "',"
            ZSql = ZSql & "'" + WCanti3 + "',"
            ZSql = ZSql & "'" + Wlote3 + "',"
            ZSql = ZSql & "'" + WLoteOri3 + "',"
            ZSql = ZSql & "'" + WTeorico + "',"
            ZSql = ZSql & "'" + WCantidadReal + "',"
            ZSql = ZSql & "'" + ZVersionI + "',"
            ZSql = ZSql & "'" + ZVersionII + "',"
            ZSql = ZSql & "'" + ZVersionIII + "')"
    
            spImpreHoja = ZSql
            Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next a
                    
            
    XLinea = ZZLugar
    For Ciclo = XLinea To 14
    
        ZZLugar = ZZLugar + 1
        WLinea = Str$(ZZLugar)
                
        WArticulo1 = ""
        WArticulo2 = ""
        WCantidad = ""
        WCanti1 = ""
        WLote1 = ""
        WLoteOri1 = ""
        WCanti2 = ""
        WLote2 = ""
        WLoteOri2 = ""
        WCanti3 = ""
        Wlote3 = ""
        WLoteOri3 = ""
        WDescripcion = ""
    
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreHoja44000 ("
        ZSql = ZSql & "Hoja ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Fecha ,"
        ZSql = ZSql & "Codigo1 ,"
        ZSql = ZSql & "Codigo2 ,"
        ZSql = ZSql & "Nombre ,"
        ZSql = ZSql & "Articulo1 ,"
        ZSql = ZSql & "Articulo2 ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "Descripcion ,"
        ZSql = ZSql & "Canti1 ,"
        ZSql = ZSql & "Lote1 ,"
        ZSql = ZSql & "LoteOri1 ,"
        ZSql = ZSql & "Canti2 ,"
        ZSql = ZSql & "Lote2 ,"
        ZSql = ZSql & "LoteOri2 ,"
        ZSql = ZSql & "Canti3 ,"
        ZSql = ZSql & "Lote3 ,"
        ZSql = ZSql & "LoteOri3 ,"
        ZSql = ZSql & "Teorico ,"
        ZSql = ZSql & "CantidadReal ,"
        ZSql = ZSql & "VersionI ,"
        ZSql = ZSql & "VersionII ,"
        ZSql = ZSql & "VersionIII )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + WHoja + "',"
        ZSql = ZSql & "'" + WLinea + "',"
        ZSql = ZSql & "'" + WFecha + "',"
        ZSql = ZSql & "'" + WCodigo1 + "',"
        ZSql = ZSql & "'" + WCodigo2 + "',"
        ZSql = ZSql & "'" + WNombre + "',"
        ZSql = ZSql & "'" + WArticulo1 + "',"
        ZSql = ZSql & "'" + WArticulo2 + "',"
        ZSql = ZSql & "'" + WCantidad + "',"
        ZSql = ZSql & "'" + WDescripcion + "',"
        ZSql = ZSql & "'" + WCanti1 + "',"
        ZSql = ZSql & "'" + WLote1 + "',"
        ZSql = ZSql & "'" + WLoteOri1 + "',"
        ZSql = ZSql & "'" + WCanti2 + "',"
        ZSql = ZSql & "'" + WLote2 + "',"
        ZSql = ZSql & "'" + WLoteOri2 + "',"
        ZSql = ZSql & "'" + WCanti3 + "',"
        ZSql = ZSql & "'" + Wlote3 + "',"
        ZSql = ZSql & "'" + WLoteOri3 + "',"
        ZSql = ZSql & "'" + WTeorico + "',"
        ZSql = ZSql & "'" + WCantidadReal + "',"
        ZSql = ZSql & "'" + ZVersionI + "',"
        ZSql = ZSql & "'" + ZVersionII + "',"
        ZSql = ZSql & "'" + ZVersionIII + "')"

        spImpreHoja = ZSql
        Set rstImpreHoja = db.OpenRecordset(spImpreHoja, dbOpenSnapshot, dbSQLPassThrough)

    Next Ciclo
End Sub







Private Sub Limpia_Click()

    Erase XEnvase

    CargaLote.Visible = False
    Erase XLote
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
    
    WEnvase1.Text = ""
    WCantiEnv1.Text = ""
    WBultos1.Text = ""
    WDescri1.Caption = ""
    WEnvase2.Text = ""
    WCantiEnv2.Text = ""
    WBultos2.Text = ""
    WDescri2.Caption = ""
    WEnvase3.Text = ""
    WCantiEnv3.Text = ""
    WBultos3.Text = ""
    WDescri3.Caption = ""
    WEnvase4.Text = ""
    WCantiEnv4.Text = ""
    WBultos4.Text = ""
    WDescri4.Caption = ""
    WEnvase5.Text = ""
    WCantiEnv5.Text = ""
    WBultos5.Text = ""
    WDescri5.Caption = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = "  /  /    "
    MarcaFactura.ListIndex = 0
    
    For a = 0 To 9
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
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
                        DBGrid1.Col = 2
                        XOriginal = Val(DBGrid1.Text)
                        DBGrid1.Col = 3
                        XCantidad = Val(DBGrid1.Text)
                        WRow = DBGrid1.Row
                        
                        ZDife = XCantidad - XOriginal
                        ZMargen = XOriginal * 0.25
                        
                        If ZDife > ZMargen Then
                        
                            m$ = "La cantidad que se desea ingresa supera en mas del 25% de la cantidad solicitada por el cliente"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                            
                                Else
                        
                            CargaLote.Visible = True
                        
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
                            
                            WEnvase1.Text = ""
                            WCantiEnv1.Text = ""
                            WBultos1.Text = ""
                            WDescri1.Caption = ""
                            WEnvase2.Text = ""
                            WCantiEnv2.Text = ""
                            WBultos2.Text = ""
                            WDescri2.Caption = ""
                            WEnvase3.Text = ""
                            WCantiEnv3.Text = ""
                            WBultos3.Text = ""
                            WDescri3.Caption = ""
                            WEnvase4.Text = ""
                            WCantiEnv4.Text = ""
                            WBultos4.Text = ""
                            WDescri4.Caption = ""
                            WEnvase5.Text = ""
                            WCantiEnv5.Text = ""
                            WBultos5.Text = ""
                            WDescri5.Caption = ""
                        
                            WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                            
                            If XLote(WLugar, 1) <> "" Then
                                WLote1.Text = XLote(WLugar, 1)
                                WCanti1.Text = XLote(WLugar, 2)
                                If Val(XLote(WLugar, 11)) <> 0 Then
                                    WEnvase1.Text = XLote(WLugar, 11)
                                    WCantiEnv1.Text = XLote(WLugar, 12)
                                    WBultos1.Text = XLote(WLugar, 21)
                                    spEnvases = "ConsultaEnvases " + "'" + WEnvase1.Text + "'"
                                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvases.RecordCount > 0 Then
                                        WDescri1.Caption = rstEnvases!Abreviatura
                                        rstEnvases.Close
                                    End If
                                End If
                            End If
                        
                            If XLote(WLugar, 3) <> "" Then
                                WLote2.Text = XLote(WLugar, 3)
                                WCanti2.Text = XLote(WLugar, 4)
                                If Val(XLote(WLugar, 13)) <> 0 Then
                                    WEnvase2.Text = XLote(WLugar, 13)
                                    WCantiEnv2.Text = XLote(WLugar, 14)
                                    WBultos2.Text = XLote(WLugar, 22)
                                    spEnvases = "ConsultaEnvases " + "'" + WEnvase2.Text + "'"
                                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvases.RecordCount > 0 Then
                                        WDescri2.Caption = rstEnvases!Abreviatura
                                        rstEnvases.Close
                                    End If
                                End If
                            End If
                        
                            If XLote(WLugar, 5) <> "" Then
                                Wlote3.Text = XLote(WLugar, 5)
                                WCanti3.Text = XLote(WLugar, 6)
                                If Val(XLote(WLugar, 15)) <> 0 Then
                                    WEnvase3.Text = XLote(WLugar, 15)
                                    WCantiEnv3.Text = XLote(WLugar, 16)
                                    WBultos3.Text = XLote(WLugar, 23)
                                    spEnvases = "ConsultaEnvases " + "'" + WEnvase3.Text + "'"
                                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvases.RecordCount > 0 Then
                                        WDescri3.Caption = rstEnvases!Abreviatura
                                        rstEnvases.Close
                                    End If
                                End If
                            End If
                        
                            If XLote(WLugar, 7) <> "" Then
                                WLote4.Text = XLote(WLugar, 7)
                                WCanti4.Text = XLote(WLugar, 8)
                                If Val(XLote(WLugar, 17)) <> 0 Then
                                    WEnvase4.Text = XLote(WLugar, 17)
                                    WCantiEnv4.Text = XLote(WLugar, 18)
                                    WBultos4.Text = XLote(WLugar, 24)
                                    spEnvases = "ConsultaEnvases " + "'" + WEnvase4.Text + "'"
                                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvases.RecordCount > 0 Then
                                        WDescri4.Caption = rstEnvases!Abreviatura
                                        rstEnvases.Close
                                    End If
                                End If
                            End If
                        
                            If XLote(WLugar, 9) <> "" Then
                                WLote5.Text = XLote(WLugar, 9)
                                WCanti5.Text = XLote(WLugar, 10)
                                If Val(XLote(WLugar, 19)) <> 0 Then
                                    WEnvase5.Text = XLote(WLugar, 19)
                                    WCantiEnv5.Text = XLote(WLugar, 20)
                                    WBultos5.Text = XLote(WLugar, 25)
                                    spEnvases = "ConsultaEnvases " + "'" + WEnvase5.Text + "'"
                                    Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvases.RecordCount > 0 Then
                                        WDescri5.Caption = rstEnvases!Abreviatura
                                        rstEnvases.Close
                                    End If
                                End If
                            End If
                        
                            WLote1.SetFocus
                            
                        End If
                        
                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
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
                    If Val(WLote1.Text) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote1.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote1.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                        ZSql = ""
                        If Val(WLote1.Text) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote1.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote1.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                    
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
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
                                
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                        "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    WSaldo1 = 0
                                End If
                                
                                aa = rstMovguia!Clave
                                rstMovguia.Close
                            End If
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                
                    End If
            End Select
                
            If WLote1.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = Wlote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    XLote(WLugar, 11) = WEnvase1.Text
                    XLote(WLugar, 12) = WCantiEnv1.Text
                    XLote(WLugar, 13) = WEnvase2.Text
                    XLote(WLugar, 14) = WCantiEnv2.Text
                    XLote(WLugar, 15) = WEnvase3.Text
                    XLote(WLugar, 16) = WCantiEnv3.Text
                    XLote(WLugar, 17) = WEnvase4.Text
                    XLote(WLugar, 18) = WCantiEnv4.Text
                    XLote(WLugar, 19) = WEnvase5.Text
                    XLote(WLugar, 20) = WCantiEnv5.Text
                    XLote(WLugar, 21) = WBultos1.Text
                    XLote(WLugar, 22) = WBultos2.Text
                    XLote(WLugar, 23) = WBultos3.Text
                    XLote(WLugar, 24) = WBultos4.Text
                    XLote(WLugar, 25) = WBultos5.Text
                    
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
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote1.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
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
    
        If Val(WEnvase1.Text) <> 0 Then
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
            
                Else
                
            WCantiEnv1.Text = ""
            WBultos1.Text = ""
            WLote2.SetFocus
                
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantiEnv1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WEmpresa) = 1 Then
            Select Case Val(WEnvase1.Text)
                Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                    WBultos1.Text = WCantiEnv1.Text
                Case Else
            End Select
        End If
        WBultos1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WBultos1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
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
                    If Val(WLote2.Text) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote2.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote2.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                    
                        ZSql = ""
                        If Val(WLote2.Text) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote2.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote2.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
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
                                
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                         "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    WSaldo2 = 0
                                End If
                                
                                rstMovguia.Close
                            End If
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
            End Select
                
            If WLote2.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = Wlote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    XLote(WLugar, 11) = WEnvase1.Text
                    XLote(WLugar, 12) = WCantiEnv1.Text
                    XLote(WLugar, 13) = WEnvase2.Text
                    XLote(WLugar, 14) = WCantiEnv2.Text
                    XLote(WLugar, 15) = WEnvase3.Text
                    XLote(WLugar, 16) = WCantiEnv3.Text
                    XLote(WLugar, 17) = WEnvase4.Text
                    XLote(WLugar, 18) = WCantiEnv4.Text
                    XLote(WLugar, 19) = WEnvase5.Text
                    XLote(WLugar, 20) = WCantiEnv5.Text
                    XLote(WLugar, 21) = WBultos1.Text
                    XLote(WLugar, 22) = WBultos2.Text
                    XLote(WLugar, 23) = WBultos3.Text
                    XLote(WLugar, 24) = WBultos4.Text
                    XLote(WLugar, 25) = WBultos5.Text
                    
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
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote2.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
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
            WEnvase2.SetFocus
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

Private Sub WEnvase2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(WEnvase2.Text) <> 0 Then
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
    
            spEnvases = "ConsultaEnvases " + "'" + WEnvase2.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WDescri2.Caption = rstEnvases!Abreviatura
                rstEnvases.Close
                WCantiEnv2.SetFocus
                    Else
                WEnvase2.SetFocus
            End If
        
            Call Conecta_Empresa
            
                Else
                
            WCantiEnv2.Text = ""
            WBultos2.Text = ""
            Wlote3.SetFocus
                
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantiEnv2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WEmpresa) = 1 Then
            Select Case Val(WEnvase2.Text)
                Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                    WBultos2.Text = WCantiEnv2.Text
                Case Else
            End Select
        End If
        WBultos2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WBultos2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
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
                    If Val(Wlote3.Text) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + Wlote3.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + Wlote3.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                        ZSql = ""
                        If Val(Wlote3.Text) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + Wlote3.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + Wlote3.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
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
                                
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                         "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    WSaldo3 = 0
                                End If
                                
                                rstMovguia.Close
                            End If
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
                    
            End Select
                
            If Wlote3.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = Wlote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    XLote(WLugar, 11) = WEnvase1.Text
                    XLote(WLugar, 12) = WCantiEnv1.Text
                    XLote(WLugar, 13) = WEnvase2.Text
                    XLote(WLugar, 14) = WCantiEnv2.Text
                    XLote(WLugar, 15) = WEnvase3.Text
                    XLote(WLugar, 16) = WCantiEnv3.Text
                    XLote(WLugar, 17) = WEnvase4.Text
                    XLote(WLugar, 18) = WCantiEnv4.Text
                    XLote(WLugar, 19) = WEnvase5.Text
                    XLote(WLugar, 20) = WCantiEnv5.Text
                    XLote(WLugar, 21) = WBultos1.Text
                    XLote(WLugar, 22) = WBultos2.Text
                    XLote(WLugar, 23) = WBultos3.Text
                    XLote(WLugar, 24) = WBultos4.Text
                    XLote(WLugar, 25) = WBultos5.Text
                    
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
                    Wlote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
                    Else
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + Wlote3.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
            End If
            
            If WEstado = "N" Then
                Wlote3.SetFocus
            End If
    
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo3 >= Val(WCanti3.Text) Then
            WCanti3.Text = Pusing("###,###.##", WCanti3.Text)
            WEnvase3.SetFocus
                Else
            XSaldo3 = WSaldo3
            XSaldo3 = Pusing("###,###.##", XSaldo3)
            m$ = XTerminado + " Cantidad Insuficiente Stock : " + XSaldo3
            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
            Wlote3.SetFocus
        End If
        Rem WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
        Rem Wlote3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WEnvase3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(WEnvase3.Text) <> 0 Then
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
    
            spEnvases = "ConsultaEnvases " + "'" + WEnvase3.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WDescri3.Caption = rstEnvases!Abreviatura
                rstEnvases.Close
                WCantiEnv3.SetFocus
                    Else
                WEnvase3.SetFocus
            End If
        
            Call Conecta_Empresa
            
                Else
                
            WCantiEnv3.Text = ""
            WBultos3.Text = ""
            WLote4.SetFocus
                
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantiEnv3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WEmpresa) = 1 Then
            Select Case Val(WEnvase3.Text)
                Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                    WBultos3.Text = WCantiEnv3.Text
                Case Else
            End Select
        End If
        WBultos3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WBultos3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
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
                    If Val(WLote4.Text) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote4.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote4.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                        ZSql = ""
                        If Val(WLote4.Text) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote4.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote4.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
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
                                
                                WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                                If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                    m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                         "Por favor comuniquese con el laboratorio para su revalida"
                                    G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                    WSaldo4 = 0
                                End If
                                
                                rstMovguia.Close
                            End If
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
                    
            End Select
                
            If WLote4.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = Wlote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    XLote(WLugar, 11) = WEnvase1.Text
                    XLote(WLugar, 12) = WCantiEnv1.Text
                    XLote(WLugar, 13) = WEnvase2.Text
                    XLote(WLugar, 14) = WCantiEnv2.Text
                    XLote(WLugar, 15) = WEnvase3.Text
                    XLote(WLugar, 16) = WCantiEnv3.Text
                    XLote(WLugar, 17) = WEnvase4.Text
                    XLote(WLugar, 18) = WCantiEnv4.Text
                    XLote(WLugar, 19) = WEnvase5.Text
                    XLote(WLugar, 20) = WCantiEnv5.Text
                    XLote(WLugar, 21) = WBultos1.Text
                    XLote(WLugar, 22) = WBultos2.Text
                    XLote(WLugar, 23) = WBultos3.Text
                    XLote(WLugar, 24) = WBultos4.Text
                    XLote(WLugar, 25) = WBultos5.Text
                    
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
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote4.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
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
            WEnvase4.SetFocus
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

Private Sub WEnvase4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(WEnvase4.Text) <> 0 Then
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
    
            spEnvases = "ConsultaEnvases " + "'" + WEnvase4.Text + "'"
            Set rstEnvases = db.OpenRecordset(spEnvases, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnvases.RecordCount > 0 Then
                WDescri4.Caption = rstEnvases!Abreviatura
                rstEnvases.Close
                WCantiEnv4.SetFocus
                    Else
                WEnvase4.SetFocus
            End If
        
            Call Conecta_Empresa
            
                Else
                
            WCantiEnv4.Text = ""
            WBultos4.Text = ""
            WLote5.SetFocus
                
        End If
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantiEnv4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WEmpresa) = 1 Then
            Select Case Val(WEnvase4.Text)
                Case 20, 21, 22, 23, 24, 25, 26, 28, 30
                    WBultos4.Text = WCantiEnv4.Text
                Case Else
            End Select
        End If
        WBultos4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WBultos4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            WEstado = ""
            WBuscaEmpresa = ""
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                
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
                    If Val(WLote5.Text) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote5.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote5.Text + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                        ZSql = ""
                        If Val(WLote5.Text) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote5.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote5.Text + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    WBuscaEmpresa = WEmpresa
                    Call Conecta_Empresa
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                            
                            WMarcaVencida = IIf(IsNull(rstHoja!MarcaVencida), "", rstHoja!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
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
                            
                            WMarcaVencida = IIf(IsNull(rstMovguia!MarcaVencida), "", rstMovguia!MarcaVencida)
                            If WMarcaVencida = "S" Or WMarcaVencida = "V" Then
                                m$ = "La Partida se encuentra vencida o ya paso mas del 75% de su vida util" + Chr$(13) + _
                                     "Por favor comuniquese con el laboratorio para su revalida"
                                G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                                WSaldo5 = 0
                            End If
                            rstMovguia.Close
                        End If
                        
                        WBuscaEmpresa = WEmpresa
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                    
                    End If
                    
            End Select
                
            If WLote5.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = DBGrid1.FirstRow + DBGrid1.Row + 1
                    XLote(WLugar, 1) = WLote1.Text
                    XLote(WLugar, 2) = WCanti1.Text
                    XLote(WLugar, 3) = WLote2.Text
                    XLote(WLugar, 4) = WCanti2.Text
                    XLote(WLugar, 5) = Wlote3.Text
                    XLote(WLugar, 6) = WCanti3.Text
                    XLote(WLugar, 7) = WLote4.Text
                    XLote(WLugar, 8) = WCanti4.Text
                    XLote(WLugar, 9) = WLote5.Text
                    XLote(WLugar, 10) = WCanti5.Text
                    XLote(WLugar, 11) = WEnvase1.Text
                    XLote(WLugar, 12) = WCantiEnv1.Text
                    XLote(WLugar, 13) = WEnvase2.Text
                    XLote(WLugar, 14) = WCantiEnv2.Text
                    XLote(WLugar, 15) = WEnvase3.Text
                    XLote(WLugar, 16) = WCantiEnv3.Text
                    XLote(WLugar, 17) = WEnvase4.Text
                    XLote(WLugar, 18) = WCantiEnv4.Text
                    XLote(WLugar, 19) = WEnvase5.Text
                    XLote(WLugar, 20) = WCantiEnv5.Text
                    XLote(WLugar, 21) = WBultos1.Text
                    XLote(WLugar, 22) = WBultos2.Text
                    XLote(WLugar, 23) = WBultos3.Text
                    XLote(WLugar, 24) = WBultos4.Text
                    XLote(WLugar, 25) = WBultos5.Text
                    
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
                If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                    Select Case Val(WBuscaEmpresa)
                        Case 1
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta I"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 11
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta VII (FARMA)"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case 7
                            m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente " + _
                                Chr$(13) + "El stock disponible se debe encontrar en la Planta V"
                            G% = MsgBox(m$, 0, "Actualizacion de Pedidos a Facturar")
                        Case Else
                    End Select
                End If
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
                XLote(WLugar, 1) = WLote1.Text
                XLote(WLugar, 2) = WCanti1.Text
                XLote(WLugar, 3) = WLote2.Text
                XLote(WLugar, 4) = WCanti2.Text
                XLote(WLugar, 5) = Wlote3.Text
                XLote(WLugar, 6) = WCanti3.Text
                XLote(WLugar, 7) = WLote4.Text
                XLote(WLugar, 8) = WCanti4.Text
                XLote(WLugar, 9) = WLote5.Text
                XLote(WLugar, 10) = WCanti5.Text
                XLote(WLugar, 11) = WEnvase1.Text
                XLote(WLugar, 12) = WCantiEnv1.Text
                XLote(WLugar, 13) = WEnvase2.Text
                XLote(WLugar, 14) = WCantiEnv2.Text
                XLote(WLugar, 15) = WEnvase3.Text
                XLote(WLugar, 16) = WCantiEnv3.Text
                XLote(WLugar, 17) = WEnvase4.Text
                XLote(WLugar, 18) = WCantiEnv4.Text
                XLote(WLugar, 19) = WEnvase5.Text
                XLote(WLugar, 20) = WCantiEnv5.Text
                XLote(WLugar, 21) = WBultos1.Text
                XLote(WLugar, 22) = WBultos2.Text
                XLote(WLugar, 23) = WBultos3.Text
                XLote(WLugar, 24) = WBultos4.Text
                XLote(WLugar, 25) = WBultos5.Text
                
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
ReDim UserData(0 To 5, 0 To 100)

mTotalRows& = 100

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
        Case 1, 3, 5, 6, 7, 10, 11
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
    Erase XLote
    
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
    
    WEnvase1.Text = ""
    WCantiEnv1.Text = ""
    WBultos1.Text = ""
    WDescri1.Caption = ""
    WEnvase2.Text = ""
    WCantiEnv2.Text = ""
    WBultos2.Text = ""
    WDescri2.Caption = ""
    WEnvase3.Text = ""
    WCantiEnv3.Text = ""
    WBultos3.Text = ""
    WDescri3.Caption = ""
    WEnvase4.Text = ""
    WCantiEnv4.Text = ""
    WBultos4.Text = ""
    WDescri4.Caption = ""
    WEnvase5.Text = ""
    WCantiEnv5.Text = ""
    WBultos5.Text = ""
    WDescri5.Caption = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    MarcaFactura.Clear
    
    MarcaFactura.AddItem ""
    MarcaFactura.AddItem "Disponible"
    
    MarcaFactura.ListIndex = 0
    
    Renglon = 0
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Pedido.SetFocus
     
End Sub

Private Sub Proceso_Click()

    Erase XEnvase
        
    For a = 0 To 9
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    
    Erase Auxiliar
    Erase ClavePedido
    Erase ZVector
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
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
                
                    Rem dada
                    Rem ojo
                    Rem ver
                    Canti = !Cantidad - !Facturado
                    Rem Canti = !Cantidad
                    
                    If Canti > 0 Then
                
                        Renglon = Renglon + 1
                
                        Lugar1 = Int((Renglon - 1) / 10) * 10
                        Lugar2 = Renglon - Lugar1
                
                        DBGrid1.FirstRow = Lugar1
                        DBGrid1.Row = Lugar2 - 1
                
                        DBGrid1.Col = 0
                        DBGrid1.Text = !Terminado
                        Auxi1 = !Terminado
                
                        Rem dada
                        Rem ojo
                        Rem ver
                        DBGrid1.Col = 2
                        DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                        Rem DBGrid1.Text = Pusing("###,###.##", Str$(!Cantidad))
                
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
                        
                        XLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "0", rstPedido!lote1)
                        XLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "0", rstPedido!CantiLote1)
                        XLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "0", rstPedido!lote2)
                        XLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "0", rstPedido!CantiLote2)
                        XLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "0", rstPedido!lote3)
                        XLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "0", rstPedido!CantiLote3)
                        XLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "0", rstPedido!lote4)
                        XLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "0", rstPedido!CantiLote4)
                        XLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "0", rstPedido!lote5)
                        XLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "0", rstPedido!CantiLote5)
                    
                        XLote(WLugar, 11) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                        XLote(WLugar, 12) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                        XLote(WLugar, 13) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env2)
                        XLote(WLugar, 14) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv2)
                        XLote(WLugar, 15) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env3)
                        XLote(WLugar, 16) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv3)
                        XLote(WLugar, 17) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env4)
                        XLote(WLugar, 18) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv4)
                        XLote(WLugar, 19) = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env5)
                        XLote(WLugar, 20) = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv5)
                        
                        XLote(WLugar, 21) = IIf(IsNull(rstPedido!Bultos1), "0", rstPedido!Bultos1)
                        XLote(WLugar, 22) = IIf(IsNull(rstPedido!Bultos2), "0", rstPedido!Bultos2)
                        XLote(WLugar, 23) = IIf(IsNull(rstPedido!Bultos3), "0", rstPedido!Bultos3)
                        XLote(WLugar, 24) = IIf(IsNull(rstPedido!Bultos4), "0", rstPedido!Bultos4)
                        XLote(WLugar, 25) = IIf(IsNull(rstPedido!Bultos5), "0", rstPedido!Bultos5)
                    
                        Rem Envase1.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env1)
                        Rem Canti1.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv1)
                        Rem Envase2.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env2)
                        Rem Canti2.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv2)
                        Rem Envase3.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env3)
                        Rem Canti3.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv3)
                        Rem Envase4.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env4)
                        Rem Canti4.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv4)
                        Rem Envase5.Text = IIf(IsNull(rstPedido!Env1), "0", rstPedido!Env5)
                        Rem Canti5.Text = IIf(IsNull(rstPedido!CantiEnv1), "0", rstPedido!CantiEnv5)
                    
                    
                        Auxiliar(Renglon, 1) = Auxi1
                        Auxiliar(Renglon, 2) = Canti
                        
                        ClavePedido(Renglon) = rstPedido!Clave
                    
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        ZVector(Renglon, 1) = !Terminado
                        ZVector(Renglon, 2) = ""
                        ZVector(Renglon, 3) = Pusing("###,###.##", Str$(!Cantidad - !Facturado))
                        ZVector(Renglon, 4) = ""
                        ZVector(Renglon, 5) = IIf(IsNull(rstPedido!Especificaciones), "0", rstPedido!Especificaciones)
                        
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
        Descri2.Caption = rstEnvases!Abreviatura
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
    
    For DA = 1 To WRenglon
    
        Renglon = Renglon + 1
    
        Auxi1 = Auxiliar(DA, 1)
        Canti = Auxiliar(DA, 2)
        
        ClavePrecios = Cliente.Text + Auxi1
        
        If Left$(Auxi1, 2) <> "PT" And Left$(Auxi1, 2) <> "PE" And Left$(Auxi1, 2) <> "YQ" And Left$(Auxi1, 2) <> "YF" And Left$(Auxi1, 2) <> "YP" And Left$(Auxi1, 2) <> "YH" Then
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
                    Case 1, 3, 5, 6, 7, 10, 11
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
                    ZVector(Renglon, 4) = Str$(Precio)
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
                    
                    ZVector(Renglon, 2) = rstArticulo!Descripcion
                
                    rstArticulo.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 9 Step 2
                    If Val(XLote(DA, Ciclo)) = 0 Then
                        XLote(DA, Ciclo) = ""
                            Else
                        ZEntra = "N"
                        
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
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + XLote(DA, Ciclo) + "'"
                        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArti + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            XLote(DA, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            ZEntra = "S"
                            rstLaudo.Close
                        End If
                        
                        If ZEntra = "N" Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Lote = " + "'" + XLote(DA, Ciclo) + "'"
                            ZSql = ZSql + " and Guia.Articulo = " + "'" + WArti + "'"
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                XLote(DA, Ciclo) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
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

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
            
            Case Else
            
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
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
                    
                    ZVector(Renglon, 2) = rstPrecios!Descripcion
                    ZVector(Renglon, 4) = Str$(Precio)
                    
                    rstPrecios.Close
                End If
                
                Call Conecta_Empresa
                
                For Ciclo = 1 To 9 Step 2
                    If Val(XLote(DA, Ciclo)) = 0 Then
                        XLote(DA, Ciclo) = ""
                    End If
                Next Ciclo

                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Canti) * Precio)
                End If
                
        End Select
        
    Next DA
    
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
            Case 1, 3, 5, 6, 7, 10, 11
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
            Rem dada
            Rem ojo
            Rem ver
            If rstPedido!Autorizo <> "X" Then
            Rem If rstPedido!Autorizo = "W" Then
                rstPedido.Close
                Call Conecta_Empresa
             
                m$ = "EL PEDIDO NO FUE AUTORIZADO"
                a% = MsgBox(m$, 0, "Actualizacion de Pedidos")
                
                    Else
                    
                WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                WMarcaFactura = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
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
                
                If Left$(rstPedido!Terminado, 4) = "PT-4" Then
                    WTipoPedido = "TA"
                End If
                
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
            Case 1, 3, 5, 6, 7, 10, 11
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
            Case 1, 3, 5, 6, 7, 10, 11
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
            Descri2.Caption = rstEnvases!Abreviatura
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
            Case 1, 3, 5, 6, 7, 10, 11
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
            Case 1, 3, 5, 6, 7, 10, 11
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
            Case 1, 3, 5, 6, 7, 10, 11
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
    If Wlote3.Text <> "" Then
        Suma = Suma + Val(WCanti3.Text)
    End If
    If WLote4.Text <> "" Then
        Suma = Suma + Val(WCanti4.Text)
    End If
    If WLote5.Text <> "" Then
        Suma = Suma + Val(WCanti5.Text)
    End If
    Rem by nan
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
            
            If Left$(XTerminado, 2) <> "PT" And Left$(XTerminado, 2) <> "PE" And Left$(XTerminado, 2) <> "YQ" And Left$(XTerminado, 2) <> "YF" And Left$(XTerminado, 2) <> "YP" And Left$(XTerminado, 2) <> "YH" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
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
                    If Val(WLote) = 0 Then
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                            Else
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Laudo.Laudo = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    End If
                    spLaudo = ZSql
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
                        ZSql = ""
                        If Val(WLote) = 0 Then
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                                Else
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        End If
                        spMovguia = ZSql
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
                    
                    Call Conecta_Empresa
                    
                    If WEntra = "N" Then
                        m$ = "Partida Inexistente"
                        a = MsgBox(m$, 0, "ACTUALIZACION DE PEDIDOS")
                        WEstado = "N"
                    End If
                
                Case Else
                    WEntra = "N"
                    WControla = 0
                    
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
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
                                
                        Call Conecta_Empresa
                        
                
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

Private Sub reImpre_Click()
    Call Impresion
End Sub

Private Sub Impresion()

    On Error GoTo WError
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
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
    
        ZZVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
        ZZMarcaFactura = IIf(IsNull(rstPedido!MarcaFactura), "0", rstPedido!MarcaFactura)
        ZZCliente = rstPedido!Cliente
        ZZFecha = rstPedido!Fecha
        ZZFecEntrega = rstPedido!FecEntrega
        ZZObservaciones = rstPedido!Observaciones
        ZZLugarDirEntrega = IIf(IsNull(rstPedido!DirEntrega), "1", rstPedido!DirEntrega)
        ZZTipoped = rstPedido!Tipoped
        ZZVia = rstPedido!Via
        ZZOrden = rstPedido!OrdenCpa
        
        Select Case rstPedido!TipoPedido
            Case 1
                ZZTipoPedido = "CO"
            Case 3
                ZZTipoPedido = "BI"
            Case 4
                ZZTipoPedido = "FA"
            Case 5
                ZZTipoPedido = "PG"
            Case Else
                ZZTipoPedido = "PT"
        End Select
            
        rstPedido.Close
        
        spCliente = "ConsultaCliente " + "'" + ZZCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZDesCliente = rstCliente!Razon
            
            ZZDirentrega = ""
            ZZDesPago = ""

            ZDirEntrega(1) = rstCliente!DirEntrega
            ZDirEntrega(2) = Trim(IIf(IsNull(rstCliente!DirEntregaII), "", rstCliente!DirEntregaII))
            ZDirEntrega(3) = Trim(IIf(IsNull(rstCliente!DirEntregaIII), "", rstCliente!DirEntregaIII))
            ZDirEntrega(4) = Trim(IIf(IsNull(rstCliente!DirEntregaIV), "", rstCliente!DirEntregaIV))
            ZDirEntrega(5) = Trim(IIf(IsNull(rstCliente!DirEntregaV), "", rstCliente!DirEntregaV))

            ZZDirentrega = ZDirEntrega(ZZLugarDirEntrega)
            ZZPago = Str$(rstCliente!Pago1)
            
            Erase WEspecif
            
            WEspecif(1) = ""
            WEspecif(2) = IIf(IsNull(rstCliente!Especif1), "", rstCliente!Especif1)
            WEspecif(3) = IIf(IsNull(rstCliente!Especif2), "", rstCliente!Especif2)
            WEspecif(4) = IIf(IsNull(rstCliente!Especif3), "", rstCliente!Especif3)
            WEspecif(5) = IIf(IsNull(rstCliente!Especif4), "", rstCliente!Especif4)
            WEspecif(6) = IIf(IsNull(rstCliente!Especif5), "", rstCliente!Especif5)
            WEspecif(7) = IIf(IsNull(rstCliente!Especif6), "", rstCliente!Especif6)
            WEspecif(8) = IIf(IsNull(rstCliente!Especif7), "", rstCliente!Especif7)
            WEspecif(9) = IIf(IsNull(rstCliente!Especif8), "", rstCliente!Especif8)
            WEspecif(10) = IIf(IsNull(rstCliente!Especif9), "", rstCliente!Especif9)
            For CicloEspecif = 1 To 10
                WEspecif(CicloEspecif) = RTrim(WEspecif(CicloEspecif))
            Next CicloEspecif
            
            rstCliente.Close
            
            spPago = "ConsultaPago " + "'" + ZZPago + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                ZZDesPago = rstPago!Nombre
                rstPago.Close
            End If
            
        End If
        
    End If
        
    Call Conecta_Empresa
    
    spImprePed = "Delete ImprePed"
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    WObservaciones = Left$(ZZObservaciones + Space$(100), 100)
    
    WTipoPedido = ""
    Select Case ZZTipoped
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
    Select Case ZZVia
        Case 1
            WVia = "Pedido de Exportacion Via : " + "Terrestre"
        Case 2
            WVia = "Pedido de Exportacion Via : " + "Maritimo"
        Case 3
            WVia = "Pedido de Exportacion Via : " + "Aereo"
        Case Else
    End Select
    
    Suma = 0
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 99
        
        Suma = Suma + 1
        Renglon = Renglon + 1
        
        WLugar = Renglon
                
        ZZArticulo = ZVector(WLugar, 1)
        ZZDescripcion = ZVector(WLugar, 2)
        ZZCantidad = ZVector(WLugar, 3)
        ZZPrecio = ZVector(WLugar, 4)
        WEspecificaciones = ZVector(WLugar, 5)
            
        If Val(ZZCantidad) <> 0 Then
        
            Erase ImpreEnvase
            LugarEnvase = 0
            
            For Cicla = 1 To 6 Step 2
                If Val(XEnvase(WLugar, Cicla)) <> 0 Then
                    LugarEnvase = LugarEnvase + 1
                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WLugar, Cicla) + "'"
                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If rstEnvase.RecordCount > 0 Then
                        WAbre = rstEnvase!Abreviatura
                        rstEnvase.Close
                            Else
                        WAbre = ""
                    End If
                    
                    ImpreEnvase(LugarEnvase) = Alinea("###", Str$(XEnvase(WLugar, Cicla + 1))) + " " + Left$(WAbre, 8)
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
            ZEmpresa = WNombreEmpresa
            ZVersion = ZZVersion
            ZCliente = Cliente.Text
            ZNombre = DesCliente.Caption
            ZFecha = Fecha.Text
            ZFechaent = ZZFecEntrega
            ZTipoPedido = WTipoPedido
            ZCondicion = ZZDesPago
            ZEntrega = ZZDirentrega
            ZObservaciones1 = Left$(ZZObservaciones, 50)
            ZObservaciones2 = Right$(ZZObservaciones, 50)
            ZOrden = ZZOrden
            ZArticulo = ZZArticulo
            ZDescripcion = ZZDescripcion
            ZPrecio = ZZPrecio
            ZCantidad = ZZCantidad
            ZEnvase = ImpreEnvase(1)
            
            spImprePed = "INSERT INTO ImprePed (" + _
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
                        "Cantidad , Envase )" + _
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
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
            Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
            
            If WEspecificaciones <> "" And WEspecificaciones <> "0" Then
            
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
                ZVersion = ZZVersion
                ZCliente = Cliente.Text
                ZNombre = DesCliente.Caption
                ZFecha = Fecha.Text
                ZFechaent = ZZFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = ZZDesPago
                ZEntrega = ZZDirentrega
                ZObservaciones1 = Left$(ZZObservaciones, 50)
                ZObservaciones2 = Right$(ZZObservaciones, 50)
                ZOrden = ZZOrden
                ZArticulo = "Especif.:"
                ZDescripcion = WEspecificaciones
                ZPrecio = "0"
                ZCantidad = "0"
                ZEnvase = ""
                
                spImprePed = "INSERT INTO ImprePed (" + _
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
                        "Cantidad , Envase )" + _
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
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
                Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
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
                ZEmpresa = WNombreEmpresa
                ZVersion = ZZVersion
                ZCliente = Cliente.Text
                ZNombre = DesCliente.Caption
                ZFecha = Fecha.Text
                ZFechaent = ZZFecEntrega
                ZTipoPedido = WTipoPedido
                ZCondicion = ZZDesPago
                ZEntrega = ZZDirentrega
                ZObservaciones1 = Left$(ZZObservaciones, 50)
                ZObservaciones2 = Right$(ZZObservaciones, 50)
                ZOrden = ZZOrden
                ZArticulo = ""
                ZDescripcion = ""
                ZPrecio = "0"
                ZCantidad = "0"
                ZEnvase = ImpreEnvase(Ciclo)
                
                spImprePed = "INSERT INTO ImprePed (" + _
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
                        "Cantidad , Envase )" + _
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
                        "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                        
                Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
                
        End If
            
    Next a
    
    SumaEspe = 0
    
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
        ZEmpresa = WNombreEmpresa
        ZVersion = ZZVersion
        ZCliente = Cliente.Text
        ZNombre = DesCliente.Caption
        ZFecha = Fecha.Text
        ZFechaent = ZZFecEntrega
        ZTipoPedido = WTipoPedido
        ZCondicion = ZZDesPago
        ZEntrega = ZZDirentrega
        ZObservaciones1 = Left$(ZZObservaciones, 50)
        ZObservaciones2 = Right$(ZZObservaciones, 50)
        ZOrden = ZZOrden
        ZArticulo = ""
        ZDescripcion = WEspecif(SumaEspe)
        ZPrecio = "0"
        ZCantidad = "0"
        ZEnvase = ""
                        
        spImprePed = "INSERT INTO ImprePed (" + _
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
                    "Cantidad , Envase )" + _
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
                    "'" + ZCantidad + "'," + "'" + ZEnvase + "')"
                                
        Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ImprePed SET "
    ZSql = ZSql + "Via = " + "'" + UCase(WVia) + "'"
    spImprePed = ZSql
    Set rstImprePed = db.OpenRecordset(spImprePed, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT ImprePed.Pedido, ImprePed.Version, ImprePed.Cliente, ImprePed.Nombre, ImprePed.Fecha, ImprePed.FechaEnt, ImprePed.Condicion, ImprePed.Entrega, ImprePed.Observaciones1, ImprePed.Observaciones2, ImprePed.Orden, ImprePed.Articulo, ImprePed.Descripcion, ImprePed.Precio, ImprePed.Cantidad, ImprePed.Envase, ImprePed.Via " _
                    + "From " _
                    + DSQ + ".dbo.ImprePed ImprePed " _
                    + "Where " _
                    + "ImprePed.Pedido >= 0 AND ImprePed.Pedido <= 999999 "
                        
    Listado.Connect = Connect()
    If ZZTipoped = 5 Or ZZTipoped = 6 Then
        Listado.ReportFileName = "ImprepedsqlMuestra.rpt"
            Else
        Listado.ReportFileName = "Imprepedsqlsp.rpt"
    End If
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

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
                Case 1, 3, 5, 6, 7, 10, 11
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
                Case 1, 3, 5, 6, 7, 10, 11
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
                    Descri2.Caption = rstEnvases!Descripcion
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




