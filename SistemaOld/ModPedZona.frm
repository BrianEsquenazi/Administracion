VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModPedZona 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Actualizacion de Pedidos a Facturar de Colorantes"
   ClientHeight    =   7830
   ClientLeft      =   15
   ClientTop       =   510
   ClientWidth     =   11910
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.CommandButton ImpreEti 
      Caption         =   "Etiquetas"
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
      TabIndex        =   56
      Top             =   120
      Width           =   1215
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   39
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
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
      TabIndex        =   36
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector2 
      Height          =   615
      Left            =   1080
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   393216
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
      TabIndex        =   34
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
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   32
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
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   24
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   9
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
         TabIndex        =   57
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox WTipo1 
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
         TabIndex        =   54
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox WTipo2 
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
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox WTipo3 
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
         TabIndex        =   52
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox WTipo4 
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
         TabIndex        =   51
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox WTipo5 
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
         TabIndex        =   50
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox WEti1 
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
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox WEti2 
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
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox WEti3 
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
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox WEti4 
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
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox WEti5 
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
         Top             =   2040
         Width           =   735
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
         TabIndex        =   35
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
         Left            =   1320
         TabIndex        =   21
         Top             =   2040
         Width           =   855
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
         TabIndex        =   20
         Top             =   1680
         Width           =   855
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   19
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   18
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
         Left            =   1320
         TabIndex        =   17
         Top             =   1320
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
         Left            =   1320
         TabIndex        =   16
         Top             =   960
         Width           =   855
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
         TabIndex        =   15
         Top             =   600
         Width           =   855
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   14
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
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
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   2880
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Etiq."
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
         Left            =   2160
         TabIndex        =   49
         Top             =   240
         Width           =   735
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
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   855
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
         Left            =   120
         TabIndex        =   10
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   4
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
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
      TabIndex        =   29
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
      Height          =   5895
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   16777152
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgModPedZona"
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
Private XLote(100, 22) As String

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
Dim XEti1 As String
Dim XTipo1 As String
Dim XEti2 As String
Dim XTipo2 As String
Dim XEti3 As String
Dim XTipo3 As String
Dim XEti4 As String
Dim XTipo4 As String
Dim XEti5 As String
Dim XTipo5 As String

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
Dim ZEti(100, 5) As String
Dim ZTipo(100, 5) As String

Dim WWLote As String
Dim WWTipo As String

Dim ZLugarDirEntrega As Integer
Dim ZDirEntrega(10) As String
Dim WEspecif(100) As String

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
    PrgModPedZona.Hide
    Unload Me
    PrgModifColor.Show
End Sub


Private Sub ConfirmaCargaLote_Click()

    Suma = 0
    
    WVector1.Col = 4
    XCantidad = Val(WVector1.Text)
    WVector1.Col = 1
    XTerminado = WVector1.Text
    
    
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
    
    Dife = XCantidad - Suma
    Pasa = ""
    
    If XCantidad <> Suma Then
        If XCantidad > Suma Then
            T$ = "Modificacion de Pedidos"
            m$ = "Articulo :" + XTerminado + Chr$(13) + "Cantidad Solicitada: " + Str$(XCantidad) + Chr$(13) + "Cantidad a entregar : " + Str$(Suma) + Chr$(13) + "Pendiente de Entrega : " + Str$(Dife) + Chr$(13) + "Confirma este Procedimiento"
            Respuesta% = MsgBox(m$, 32 + 4 + 256, T$)
            If Respuesta% = 6 Then
                Pasa = "S"
            End If
                Else
            Pasa = "S"
        End If
            Else
        Pasa = "S"
    End If
    
    If Pasa = "S" Then
        WLugar = WVector1.Row
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
        XLote(WLugar, 13) = WEti1.Text
        XLote(WLugar, 14) = WEti2.Text
        XLote(WLugar, 15) = WEti3.Text
        XLote(WLugar, 16) = WEti4.Text
        XLote(WLugar, 17) = WEti5.Text
        XLote(WLugar, 18) = WTipo1.Text
        XLote(WLugar, 19) = WTipo2.Text
        XLote(WLugar, 20) = WTipo3.Text
        XLote(WLugar, 21) = WTipo4.Text
        XLote(WLugar, 22) = WTipo5.Text
        CargaLote.Visible = False
        If WVector1.Row < 40 Then
            WVector1.Row = WVector1.Row + 1
            WRow = WVector1.Row
            XRow = WVector1.Row
            WVector1.Col = 4
        End If
        WVector1.Row = XRow
        WVector1.Col = 3
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_ImprePedCol
    OPEN_FILE_ImpreEtiDy
End Sub

Private Sub ImpreEti_Click()
    
    Da = 0
    With rstImpreEtiDy
        .Index = "Renglon"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                     Exit Do
                End If
            Loop
        End If
    End With
    
    WLugar = 0
    Xlugar = 0
    RowIni = WVector1.Row
    Rowfin = WVector1.RowSel
    
    For a = RowIni To Rowfin
        
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
                    
                Xlugar = a
                
                XLote1 = ""
                XLote2 = ""
                XLote3 = ""
                XLote4 = ""
                XLote5 = ""
                
                XLote1 = XLote(Xlugar, 1)
                XLote2 = XLote(Xlugar, 3)
                XLote3 = XLote(Xlugar, 5)
                XLote4 = XLote(Xlugar, 7)
                XLote5 = XLote(Xlugar, 9)
                
                XCantiLote1 = XLote(Xlugar, 2)
                XCantiLote2 = XLote(Xlugar, 4)
                XCantiLote3 = XLote(Xlugar, 6)
                XCantiLote4 = XLote(Xlugar, 8)
                XCantiLote5 = XLote(Xlugar, 10)
                XEti1 = XLote(Xlugar, 13)
                XEti2 = XLote(Xlugar, 14)
                XEti3 = XLote(Xlugar, 15)
                XEti4 = XLote(Xlugar, 16)
                XEti5 = XLote(Xlugar, 17)
                XTipo1 = XLote(Xlugar, 18)
                XTipo2 = XLote(Xlugar, 19)
                XTipo3 = XLote(Xlugar, 20)
                XTipo4 = XLote(Xlugar, 21)
                XTipo5 = XLote(Xlugar, 22)
                
                WLugar = WLugar + 1
                        
                ZLote(WLugar, 1) = XLote1
                ZLote(WLugar, 2) = XLote2
                ZLote(WLugar, 3) = XLote3
                ZLote(WLugar, 4) = XLote4
                ZLote(WLugar, 5) = XLote5
                            
                ZCanti(WLugar, 1) = XCantiLote1
                ZCanti(WLugar, 2) = XCantiLote2
                ZCanti(WLugar, 3) = XCantiLote3
                ZCanti(WLugar, 4) = XCantiLote4
                ZCanti(WLugar, 5) = XCantiLote5
                            
                ZEti(WLugar, 1) = XEti1
                ZEti(WLugar, 2) = XEti2
                ZEti(WLugar, 3) = XEti3
                ZEti(WLugar, 4) = XEti4
                ZEti(WLugar, 5) = XEti5
                            
                ZTipo(WLugar, 1) = XTipo1
                ZTipo(WLugar, 2) = XTipo2
                ZTipo(WLugar, 3) = XTipo3
                ZTipo(WLugar, 4) = XTipo4
                ZTipo(WLugar, 5) = XTipo5
                            
                ZLoteII(WLugar, 1) = XPedido
                ZLoteII(WLugar, 2) = Articulo
                ZLoteII(WLugar, 3) = Cliente.Text
                
            End If
                    
        End If
            
    Next a
    
    Renglon = 0
    
    For Da = 1 To WLugar
    
        WPedido = ZLoteII(Da, 1)
        WTerminado = ZLoteII(Da, 2)
        WCliente = ZLoteII(Da, 3)
        
        If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "DQ" Then
            WTipopro = "M"
                Else
            WTipopro = "T"
        End If
        
        WDescripcion = ""
        WRazon = ""
        
        spCliente = "ConsultaCliente " + "'" + WCliente + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WRazon = rstCliente!Razon
            rstCliente.Close
        End If
        
        Select Case WTipopro
            Case "M"
                WArti = Left$(WTerminado, 3) + Right$(WTerminado, 7)
                
                spArticulo = "ConsultaArticulo " + "'" + WArti + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
            Case Else
                spPrecios = "ConsultaPrecios " + "'" + WCliente + WTerminado + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrecios.RecordCount > 0 Then
                    WDescripcion = rstPrecios!Descripcion
                    rstPrecios.Close
                End If
                
        End Select
        
        For Ciclo1 = 1 To 5
            If Val(ZCanti(Da, Ciclo1)) <> 0 Then
                WHasta = Val(ZEti(Da, Ciclo1))
                WTipoeti = ZTipo(Da, Ciclo1)
                For Ciclo2 = 1 To WHasta
                    Renglon = Renglon + 1
                    If WTipoeti = "T" Then
                        With rstImpreEtiDy
                            .Index = "Renglon"
                            .AddNew
                            !Renglon = Renglon
                            !Pedido = WPedido
                            If Left$(WTerminado, 2) = "DY" Or Left$(WTerminado, 2) = "DS" Or Left$(WTerminado, 2) = "DQ" Then
                                !Codigo = Left$(Mid$(WTerminado, 6, 3) + Right$(WTerminado, 3) + ZLote(Da, Ciclo1), 20)
                                    Else
                                !Codigo = Left$(Mid$(WTerminado, 4, 5) + Right$(WTerminado, 3) + ZLote(Da, Ciclo1), 20)
                            End If
                            !Cliente = WCliente
                            !Descripcion = WDescripcion
                            !Razon = WRazon
                            !Lote = 0
                            !lote1 = ZLote(Da, Ciclo1)
                            !Cantidad = Val(ZCanti(Da, Ciclo1)) / Val(ZEti(Da, Ciclo1))
                            .Update
                        End With
                            Else
                        With rstImpreEtiDy
                            .Index = "Renglon"
                            .AddNew
                            !Renglon = Renglon
                            !Pedido = WPedido
                            !Codigo = ""
                            !Cliente = WCliente
                            !Descripcion = ""
                            !Razon = WRazon
                            !Lote = 0
                            !lote1 = ""
                            !Cantidad = 0
                            .Update
                        End With
                    End If
                Next Ciclo2
            End If
        Next Ciclo1
        
    Next Da

    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.ReportFileName = "WImpreEtiDy.rpt"
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.PrinterCopies = 1
    Listado.Action = 1

End Sub

Private Sub Wlote1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote1.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo1 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote1.Text + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                
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
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        XParam = "'" + WLote1.Text + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo1 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            WEntra = "S"
                            rstHoja.Close
                        End If
                
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + WLote1.Text + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo1 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
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
                    WLugar = WVector1.Row
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
                    XLote(WLugar, 13) = WEti1.Text
                    XLote(WLugar, 14) = WEti2.Text
                    XLote(WLugar, 15) = WEti3.Text
                    XLote(WLugar, 16) = WEti4.Text
                    XLote(WLugar, 17) = WEti5.Text
                    XLote(WLugar, 18) = WTipo1.Text
                    XLote(WLugar, 19) = WTipo2.Text
                    XLote(WLugar, 20) = WTipo3.Text
                    XLote(WLugar, 21) = WTipo4.Text
                    XLote(WLugar, 22) = WTipo5.Text
                    CargaLote.Visible = False
                    If WVector1.Row < 40 Then
                        WVector1.Row = WVector1.Row + 1
                        WRow = WVector1.Row
                        XRow = WVector1.Row
                        WVector1.Col = 4
                    End If
                    WVector1.Row = XRow
                    WVector1.Col = 3
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
                G% = MsgBox(m$, 0, "Emision de facturas")
            End If
    
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo1 >= Val(WCanti1.Text) Then
            WCanti1.Text = Pusing("###,###.##", WCanti1.Text)
            WEti1.SetFocus
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

Private Sub WEti1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote2.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote2.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo2 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote2.Text + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                
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
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        XParam = "'" + WLote2.Text + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo2 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            WEntra = "S"
                            rstHoja.Close
                        End If
                    
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + WLote2.Text + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo2 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
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
                    WLugar = WVector1.Row
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
                    XLote(WLugar, 13) = WEti1.Text
                    XLote(WLugar, 14) = WEti2.Text
                    XLote(WLugar, 15) = WEti3.Text
                    XLote(WLugar, 16) = WEti4.Text
                    XLote(WLugar, 17) = WEti5.Text
                    XLote(WLugar, 18) = WTipo1.Text
                    XLote(WLugar, 19) = WTipo2.Text
                    XLote(WLugar, 20) = WTipo3.Text
                    XLote(WLugar, 21) = WTipo4.Text
                    XLote(WLugar, 22) = WTipo5.Text
                    CargaLote.Visible = False
                    If WVector1.Row < 40 Then
                       WVector1.Row = WVector1.Row + 1
                       WRow = WVector1.Row
                       XRow = WVector1.Row
                       WVector1.Col = 4
                    End If
                    WVector1.Row = XRow
                    WVector1.Col = 3
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
                G% = MsgBox(m$, 0, "Emision de Facturas")
            End If
    
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCanti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WSaldo2 >= Val(WCanti2.Text) Then
            WCanti2.Text = Pusing("###,###.##", WCanti2.Text)
            WEti2.SetFocus
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

Private Sub WEti2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Wlote3.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + Wlote3.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo3 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + Wlote3.Text + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                
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
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        XParam = "'" + Wlote3.Text + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo3 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            WEntra = "S"
                            rstHoja.Close
                        End If
                        
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + Wlote3.Text + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo3 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End If
                        End If
                        
                        Call Conecta_Empresa
                                
                            Else
                    
                        WEntra = "S"
                
                    End If
                    
            End Select
                
            If Wlote3.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = WVector1.Row
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
                    XLote(WLugar, 13) = WEti1.Text
                    XLote(WLugar, 14) = WEti2.Text
                    XLote(WLugar, 15) = WEti3.Text
                    XLote(WLugar, 16) = WEti4.Text
                    XLote(WLugar, 17) = WEti5.Text
                    XLote(WLugar, 18) = WTipo1.Text
                    XLote(WLugar, 19) = WTipo2.Text
                    XLote(WLugar, 20) = WTipo3.Text
                    XLote(WLugar, 21) = WTipo4.Text
                    XLote(WLugar, 22) = WTipo5.Text
                    CargaLote.Visible = False
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
                    Wlote3.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti3.SetFocus
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
            WEti3.SetFocus
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

Private Sub WEti3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote4.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Wlote4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote4.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo4 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote4.Text + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo4 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                
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
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        XParam = "'" + WLote4.Text + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo4 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            WEntra = "S"
                            rstHoja.Close
                        End If
                
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + WLote4.Text + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo4 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
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
                    WLugar = WVector1.Row
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
                    XLote(WLugar, 13) = WEti1.Text
                    XLote(WLugar, 14) = WEti2.Text
                    XLote(WLugar, 15) = WEti3.Text
                    XLote(WLugar, 16) = WEti4.Text
                    XLote(WLugar, 17) = WEti5.Text
                    XLote(WLugar, 18) = WTipo1.Text
                    XLote(WLugar, 19) = WTipo2.Text
                    XLote(WLugar, 20) = WTipo3.Text
                    XLote(WLugar, 21) = WTipo4.Text
                    XLote(WLugar, 22) = WTipo5.Text
                    CargaLote.Visible = False
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
                    WLote4.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti4.SetFocus
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
            WEti4.SetFocus
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

Private Sub WEti4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote5.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Wlote5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                    
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote4.Text + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
                    spLaudo = ZSql
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstLaudo.RecordCount > 0 Then
                        With rstLaudo
                            .MoveFirst
                            WSaldo5 = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                            WEntra = "S"
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote5.Text + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                        spMovguia = ZSql
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        If rstMovguia.RecordCount > 0 Then
                            With rstMovguia
                                .MoveFirst
                                WSaldo5 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End With
                        End If
                        
                    End If
                    
                    Call Conecta_Empresa
                
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
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                        XParam = "'" + WLote5.Text + "','" _
                                + XTerminado + "'"
                        spHoja = "ListaHojaProducto " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHoja.RecordCount > 0 Then
                            WSaldo5 = IIf(IsNull(rstHoja!Saldo), "0", rstHoja!Saldo)
                            WEntra = "S"
                            rstHoja.Close
                        End If
                        
                        If WEntra = "N" Then
                            XParam = "'" + XTerminado + "','" _
                                    + WLote5.Text + "'"
                            spMovguia = "ListaMovguiaLote1 " + XParam
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                WSaldo5 = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                                WEntra = "S"
                                rstMovguia.Close
                            End If
                        End If
                        
                        Call Conecta_Empresa
                
                            Else
                    
                        WEntra = "S"
                    
                    End If
                    
            End Select
                
            If WLote5.Text = "" Then
                Call Verifica_Lote
                If WEstado = "S" Then
                    WLugar = WVector1.Row
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
                    XLote(WLugar, 13) = WEti1.Text
                    XLote(WLugar, 14) = WEti2.Text
                    XLote(WLugar, 15) = WEti3.Text
                    XLote(WLugar, 16) = WEti4.Text
                    XLote(WLugar, 17) = WEti5.Text
                    XLote(WLugar, 18) = WTipo1.Text
                    XLote(WLugar, 19) = WTipo2.Text
                    XLote(WLugar, 20) = WTipo3.Text
                    XLote(WLugar, 21) = WTipo4.Text
                    XLote(WLugar, 22) = WTipo5.Text
                    CargaLote.Visible = False
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
                    WLote5.SetFocus
                    Exit Sub
                End If
            End If
                
            If WEntra = "S" Then
                WCanti5.SetFocus
                    Else
                m$ = XTerminado + " Producto inexistente o Lote nro. " + WLote5.Text + " inexistente"
                G% = MsgBox(m$, 0, "Modificacion de Hoja de Produccion")
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
                WLugar = WVector1.Row
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
                XLote(WLugar, 13) = WEti1.Text
                XLote(WLugar, 14) = WEti2.Text
                XLote(WLugar, 15) = WEti3.Text
                XLote(WLugar, 16) = WEti4.Text
                XLote(WLugar, 17) = WEti5.Text
                XLote(WLugar, 18) = WTipo1.Text
                XLote(WLugar, 19) = WTipo2.Text
                XLote(WLugar, 20) = WTipo3.Text
                XLote(WLugar, 21) = WTipo4.Text
                XLote(WLugar, 22) = WTipo5.Text
                CargaLote.Visible = False
                If WVector1.Row < 40 Then
                    WVector1.Row = WVector1.Row + 1
                    WRow = WVector1.Row
                    XRow = WVector1.Row
                    WVector1.Col = 4
                End If
                WVector1.Row = XRow
                WVector1.Col = 4
                Exit Sub
            End If
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

Private Sub WEti5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTipo5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WTipo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WLote1.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
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
    WEti1.Text = ""
    WTipo1.Text = ""
    WEti2.Text = ""
    WTipo2.Text = ""
    WEti3.Text = ""
    WTipo3.Text = ""
    WEti4.Text = ""
    WTipo4.Text = ""
    WEti5.Text = ""
    WTipo5.Text = ""

    Pedido.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = 0
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
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
                        
                        XLote(WLugar, 1) = IIf(IsNull(rstPedido!lote1), "", rstPedido!lote1)
                        XLote(WLugar, 2) = IIf(IsNull(rstPedido!CantiLote1), "", rstPedido!CantiLote1)
                        XLote(WLugar, 3) = IIf(IsNull(rstPedido!lote2), "", rstPedido!lote2)
                        XLote(WLugar, 4) = IIf(IsNull(rstPedido!CantiLote2), "", rstPedido!CantiLote2)
                        XLote(WLugar, 5) = IIf(IsNull(rstPedido!lote3), "", rstPedido!lote3)
                        XLote(WLugar, 6) = IIf(IsNull(rstPedido!CantiLote3), "", rstPedido!CantiLote3)
                        XLote(WLugar, 7) = IIf(IsNull(rstPedido!lote4), "", rstPedido!lote4)
                        XLote(WLugar, 8) = IIf(IsNull(rstPedido!CantiLote4), "", rstPedido!CantiLote4)
                        XLote(WLugar, 9) = IIf(IsNull(rstPedido!lote5), "", rstPedido!lote5)
                        XLote(WLugar, 10) = IIf(IsNull(rstPedido!CantiLote5), "", rstPedido!CantiLote5)
                        XLote(WLugar, 13) = IIf(IsNull(rstPedido!Eti1), "", rstPedido!Eti1)
                        XLote(WLugar, 14) = IIf(IsNull(rstPedido!Eti2), "", rstPedido!Eti2)
                        XLote(WLugar, 15) = IIf(IsNull(rstPedido!Eti3), "", rstPedido!Eti3)
                        XLote(WLugar, 16) = IIf(IsNull(rstPedido!Eti4), "", rstPedido!Eti4)
                        XLote(WLugar, 17) = IIf(IsNull(rstPedido!Eti5), "", rstPedido!Eti5)
                        XLote(WLugar, 18) = IIf(IsNull(rstPedido!tipo1), "", rstPedido!tipo1)
                        XLote(WLugar, 19) = IIf(IsNull(rstPedido!tipo2), "", rstPedido!tipo2)
                        XLote(WLugar, 20) = IIf(IsNull(rstPedido!tipo3), "", rstPedido!tipo3)
                        XLote(WLugar, 21) = IIf(IsNull(rstPedido!tipo4), "", rstPedido!tipo4)
                        XLote(WLugar, 22) = IIf(IsNull(rstPedido!tipo5), "", rstPedido!tipo5)
                        
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
        
        If Left$(Auxi1, 2) = "DY" Or Left$(Auxi1, 2) = "DS" Or Left$(Auxi1, 2) = "DQ" Then
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
                    If Val(XLote(Da, Ciclo)) = 0 Then
                        XLote(Da, Ciclo) = ""
                            Else
                        ZEntra = "N"
                        
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Laudo"
                        ZSql = ZSql + " Where Laudo.Laudo = " + "'" + XLote(Da, Ciclo) + "'"
                        ZSql = ZSql + " and Laudo.Articulo = " + "'" + WArti + "'"
                        spLaudo = ZSql
                        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstLaudo.RecordCount > 0 Then
                            XLote(Da, Ciclo) = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                            ZEntra = "S"
                            rstLaudo.Close
                        End If
                        
                        If ZEntra = "N" Then
                        
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Guia"
                            ZSql = ZSql + " Where Guia.Lote = " + "'" + XLote(Da, Ciclo) + "'"
                            ZSql = ZSql + " and Guia.Articulo = " + "'" + WArti + "'"
                            ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
                            spMovguia = ZSql
                            Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                            If rstMovguia.RecordCount > 0 Then
                                XLote(Da, Ciclo) = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
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
                    If Val(XLote(Da, Ciclo)) = 0 Then
                        XLote(Da, Ciclo) = ""
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
                        If ControlLote(Ciclo1, 1) = ControlLote(Ciclo2, 1) <> 0 Then
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
            
            If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                WTipopro = "M"
                    Else
                WTipopro = "T"
            End If
            
            Select Case WTipopro
                Case "M"
                    WArti = Left$(XTerminado, 3) + Right$(XTerminado, 7)
                    WEntra = "N"
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Laudo"
                    ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArti + "'"
                    ZSql = ZSql + " and Laudo.PartiOri = " + "'" + WLote + "'"
                    ZSql = ZSql + " Order by Laudo.Fechaord, Laudo.Laudo"
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
                            rstLaudo.Close
                        End With
                    End If
                        
                    If WEntra = "N" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Guia"
                        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArti + "'"
                        ZSql = ZSql + " and Guia.PartiOri = " + "'" + WLote + "'"
                        ZSql = ZSql + " Order by Guia.Fechaord, Guia.Codigo"
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
                    spTerminado = "ConsultaTerminado " + "'" + XTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WControla = IIf(IsNull(rstTerminado!Controla), "0", rstTerminado!Controla)
                        rstTerminado.Close
                    End If
            
                    If WControla = 0 Then
                    
                        XEmpresa = WEmpresa
    
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
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
                WEti1.Text = ""
                WTipo1.Text = ""
                WEti2.Text = ""
                WTipo2.Text = ""
                WEti3.Text = ""
                WTipo3.Text = ""
                WEti4.Text = ""
                WTipo4.Text = ""
                WEti5.Text = ""
                WTipo5.Text = ""
                            
                WLugar = WVector1.Row
                
                If Left$(XTerminado, 2) = "DY" Or Left$(XTerminado, 2) = "DS" Or Left$(XTerminado, 2) = "DQ" Then
                                
                    If XLote(WLugar, 1) <> "" Then
                        WLote1.Text = XLote(WLugar, 1)
                        WCanti1.Text = XLote(WLugar, 2)
                        WEti1.Text = XLote(WLugar, 13)
                        WTipo1.Text = XLote(WLugar, 18)
                    End If
                    If XLote(WLugar, 3) <> "" Then
                        WLote2.Text = XLote(WLugar, 3)
                        WCanti2.Text = XLote(WLugar, 4)
                        WEti2.Text = XLote(WLugar, 14)
                        WTipo2.Text = XLote(WLugar, 19)
                    End If
                    If XLote(WLugar, 5) <> "" Then
                        Wlote3.Text = XLote(WLugar, 5)
                        WCanti3.Text = XLote(WLugar, 6)
                        WEti3.Text = XLote(WLugar, 15)
                        WTipo3.Text = XLote(WLugar, 20)
                    End If
                    If XLote(WLugar, 7) <> "" Then
                        WLote4.Text = XLote(WLugar, 7)
                        WCanti4.Text = XLote(WLugar, 8)
                        WEti4.Text = XLote(WLugar, 16)
                        WTipo4.Text = XLote(WLugar, 21)
                    End If
                    If XLote(WLugar, 9) <> "" Then
                        WLote5.Text = XLote(WLugar, 9)
                        WCanti5.Text = XLote(WLugar, 10)
                        WEti5.Text = XLote(WLugar, 17)
                        WTipo5.Text = XLote(WLugar, 22)
                    End If
                    
                        Else
                    
                    If Val(XLote(WLugar, 1)) <> 0 Then
                        WLote1.Text = XLote(WLugar, 1)
                        WCanti1.Text = XLote(WLugar, 2)
                        WEti1.Text = XLote(WLugar, 13)
                        WTipo1.Text = XLote(WLugar, 18)
                    End If
                    If Val(XLote(WLugar, 3)) <> 0 Then
                        WLote2.Text = XLote(WLugar, 3)
                        WCanti2.Text = XLote(WLugar, 4)
                        WEti2.Text = XLote(WLugar, 14)
                        WTipo2.Text = XLote(WLugar, 19)
                    End If
                    If Val(XLote(WLugar, 5)) <> 0 Then
                        Wlote3.Text = XLote(WLugar, 5)
                        WCanti3.Text = XLote(WLugar, 6)
                        WEti3.Text = XLote(WLugar, 15)
                        WTipo3.Text = XLote(WLugar, 20)
                    End If
                    If Val(XLote(WLugar, 7)) <> 0 Then
                        WLote4.Text = XLote(WLugar, 7)
                        WCanti4.Text = XLote(WLugar, 8)
                        WEti4.Text = XLote(WLugar, 16)
                        WTipo4.Text = XLote(WLugar, 21)
                    End If
                    If Val(XLote(WLugar, 9)) <> 0 Then
                        WLote5.Text = XLote(WLugar, 9)
                        WCanti5.Text = XLote(WLugar, 10)
                        WEti5.Text = XLote(WLugar, 17)
                        WTipo5.Text = XLote(WLugar, 22)
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
    Rem modificar el tama�o de las celdas
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

Private Sub Limpia_Vector2()

    WVector2.Height = 4095
    WVector2.Left = 120
    WVector2.Top = 1350
    WVector2.Width = 12000

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 12
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
                WVector2.ColWidth(Ciclo) = 800
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
                WVector2.Text = "Part.Orig"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Envase"
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Cant.Ped."
                WVector2.ColWidth(Ciclo) = 1100
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 11
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
        WTitulo(Ciclo).Visible = True
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
    Rem modificar el tama�o de las celdas
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
    WVector2.Cols = 12
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
                WVector2.ColWidth(Ciclo) = 800
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
                WVector2.Text = "Part.Orig"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 8
                WVector2.Text = "Partida"
                WVector2.ColWidth(Ciclo) = 900
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Envase"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Nro. Transito"
                WVector2.ColWidth(Ciclo) = 2000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 11
                WVector2.Text = "Stock"
                WVector2.ColWidth(Ciclo) = 800
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
        WTitulo(Ciclo).Visible = True
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
    Rem modificar el tama�o de las celdas
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

Private Sub CancelaCargaLote_Click()
    CargaLote.Visible = False
End Sub

Private Sub WLote1_DblClick()
    WProceso = 1
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    If Left$(WTerminado, 2) <> "DY" And Left$(WTerminado, 2) <> "DS" And Left$(WTerminado, 2) <> "DQ" Then
        Call ficha_Pt
            Else
        Call ficha_Mp
    End If
End Sub

Private Sub WLote2_DblClick()
    WProceso = 2
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    If Left$(WTerminado, 2) <> "DY" And Left$(WTerminado, 2) <> "DS" And Left$(WTerminado, 2) <> "DQ" Then
        Call ficha_Pt
            Else
        Call ficha_Mp
    End If
End Sub

Private Sub WLote3_DblClick()
    WProceso = 3
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    If Left$(WTerminado, 2) <> "DY" And Left$(WTerminado, 2) <> "DS" And Left$(WTerminado, 2) <> "DQ" Then
        Call ficha_Pt
            Else
        Call ficha_Mp
    End If
End Sub

Private Sub WLote4_DblClick()
    WProceso = 4
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    If Left$(WTerminado, 2) <> "DY" And Left$(WTerminado, 2) <> "DS" And Left$(WTerminado, 2) <> "DQ" Then
        Call ficha_Pt
            Else
        Call ficha_Mp
    End If
End Sub

Private Sub WLote5_DblClick()
    WProceso = 5
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    If Left$(WTerminado, 2) <> "DY" And Left$(WTerminado, 2) <> "DS" And Left$(WTerminado, 2) <> "DQ" Then
        Call ficha_Pt
            Else
        Call ficha_Mp
    End If
End Sub

Private Sub ficha_Mp()
    
    XEmpresa = WEmpresa
    
    WEmpresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Call Limpia_Vector2II
    WArticulo = Left$(WVector1.TextMatrix(WVector1.Row, 1), 3) + Right$(WVector1.TextMatrix(WVector1.Row, 1), 7)
    
    XRenglon = 0
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
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
                
                ZSaldoTransito = IIf(IsNull(rstLaudo!SaldoTransito), "0", rstLaudo!SaldoTransito)
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 And ZSaldoTransito <> 0 Then
                
                        Else
                    
                    If rstLaudo!Articulo = WArticulo Then
                    
                        ZArticulo = rstLaudo!Articulo
                        ZCantidad = rstLaudo!Liberada
                        ZFecha = rstLaudo!Fecha
                        ZLaudo = rstLaudo!Laudo
                        ZOrden = rstLaudo!Orden
                        Zdevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                        ZRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                        ZSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        ZLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                        ZPArtiOri = IIf(IsNull(rstLaudo!PartiOri), "0", rstLaudo!PartiOri)
                        ZTransito = IIf(IsNull(rstLaudo!Transito), "", rstLaudo!Transito)
                        ZSaldoTransito = IIf(IsNull(rstLaudo!SaldoTransito), "0", rstLaudo!SaldoTransito)
                        Call Redondeo(ZSaldo)
                        
                        If ZLiberada <> 0 And ZSaldo <> 0 Then
                        
                            XRenglon = XRenglon + 1
                            WVector2.Row = XRenglon
                
                            WVector2.Col = 1
                            WVector2.Text = "Laudo"
                        
                            WVector2.Col = 2
                            WVector2.Text = ZLaudo
                                               
                            WVector2.Col = 3
                            WVector2.Text = ZFecha
                        
                            WVector2.Col = 4
                            WVector2.Text = ZOrden
                        
                            WVector2.Col = 5
                            WVector2.Text = ZCantidad
                
                            WVector2.Col = 6
                            WVector2.Text = ZSaldo
                            
                            WVector2.Col = 7
                            WVector2.Text = ZPArtiOri
                
                            WVector2.Col = 8
                            WVector2.Text = ZLaudo
                            
                            WVector2.Col = 9
                            WVector2.Text = ""
                            
                            WVector2.Col = 10
                            WVector2.Text = ZTransito
                            
                            WVector2.Col = 11
                            WVector2.Text = ZSaldoTransito
                            
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
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
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
                        
                    If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                    
                        ZArticulo = rstMovguia!Articulo
                        ZCantidad = rstMovguia!Cantidad
                        ZFecha = rstMovguia!Fecha
                        ZCodigo = rstMovguia!Codigo
                        ZMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        ZTipomov = rstMovguia!Tipomov
                        ZPartida = IIf(IsNull(rstMovguia!Lote), "o", rstMovguia!Lote)
                        ZPartidaOri = IIf(IsNull(rstMovguia!PartiOri), "", rstMovguia!PartiOri)
                        ZFecha = rstMovguia!Fecha
                        If Val(ZCodigo) > 900000 Then
                            WWTipo = "Prestamo"
                            ZCodigo = ZCodigo - 900000
                                Else
                            WWTipo = "Guia In"
                        End If
                        ZSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(ZSaldo)
                                
                        If rstMovguia!Movi = "E" And ZSaldo <> 0 Then
                            
                            XRenglon = XRenglon + 1
                            WVector2.Row = XRenglon
                
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
                            WVector2.Text = ZPartidaOri
                            
                            WVector2.Col = 8
                            WVector2.Text = ZPartida
                            
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
        rstMovguia.Close
    End If
    
    For Ciclo = 1 To XRenglon
    
        XParam = "'" + WVector2.TextMatrix(Ciclo, 4) + "','" _
                 + WArticulo + "'"
        spInforme = "ListaInformeOrdenArticulo " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            WEnvase = Str$(rstInforme!Envase)
            rstInforme.Close
        End If
        
        spEnvase = "ConsultaEnvases " + "'" + WEnvase + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WVector2.TextMatrix(Ciclo, 9) = rstEnvase!Abreviatura
            rstEnvase.Close
        End If
        
        WLote = WVector2.TextMatrix(Ciclo, 7)
        If Left$(WArticulo, 2) = "DY" Then
            WTermi = "DY-00" + Mid$(WArticulo, 4, 7)
                Else
            If Left$(WArticulo, 2) = "DS" Then
                WTermi = "DS-00" + Mid$(WArticulo, 4, 7)
                    Else
                WTermi = "DQ-00" + Mid$(WArticulo, 4, 7)
            End If
        End If
        
        XParam = "'" + WLote + "'"
        spLaudo = "ListaLaudoPartiOri " + XParam
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            WLote = IIf(IsNull(rstLaudo!Laudo), "", rstLaudo!Laudo)
            rstLaudo.Close
        End If
        
    Next Ciclo
    
    WVector2.Col = 1
    WVector2.Row = 1
    
    WVector2.TopRow = 1
    
    Call Conecta_Empresa
    
End Sub


Private Sub ficha_Pt()

    XEmpresa = WEmpresa
    
    WEmpresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)


    Call Limpia_Vector2
    WTerminado = WVector1.TextMatrix(WVector1.Row, 1)
    XRenglon = 0
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
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
                        WVector2.Row = XRenglon
                
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
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
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
                        WWTipo = "Prestamo"
                        ZCodigo = WCodigo - 900000
                            Else
                        WWTipo = "Guia In"
                    End If
                    
                    If ZMovi = "E" And ZSaldo <> 0 Then
                    
                        XRenglon = XRenglon + 1
                        WVector2.Row = XRenglon
                
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
    
    
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
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
    
    Call Conecta_Empresa
    
End Sub



