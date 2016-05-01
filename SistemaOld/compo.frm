VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompo 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Composicion de Productos Terminados"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8085
   ScaleWidth      =   11880
   Visible         =   0   'False
   Begin VB.Frame PantaExplosion 
      Height          =   1215
      Left            =   720
      TabIndex        =   76
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton CierraPanta 
         Caption         =   "Cierra Pantalla"
         Height          =   615
         Left            =   3600
         TabIndex        =   79
         Top             =   4320
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid WGrillaExplo 
         Height          =   3855
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6800
         _Version        =   327680
         BackColor       =   16777152
      End
   End
   Begin VB.CommandButton Explosion 
      Caption         =   "Explosion"
      Height          =   495
      Left            =   9720
      TabIndex        =   78
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CheckBox Restriccion 
      Caption         =   "Restriccion"
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
      Left            =   9720
      TabIndex        =   75
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8520
      TabIndex        =   74
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CaratulaPDF 
      Caption         =   "Caratula PDF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9720
      TabIndex        =   73
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   3720
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   2895
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
      Index           =   5
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   3360
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   3240
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   3360
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
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   3360
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame5 
      Height          =   3855
      Left            =   9600
      TabIndex        =   46
      Top             =   2040
      Width           =   2295
      Begin VB.TextBox Vida 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   65
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Embalaje 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   55
         Top             =   555
         Width           =   975
      End
      Begin VB.TextBox Naciones 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Intervencion 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   53
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox EstadoProducto 
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ComboBox Carga 
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Riesgo 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   50
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Secundario 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   49
         Top             =   1150
         Width           =   735
      End
      Begin VB.TextBox Clase 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   48
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Caracteristicas 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   47
         Text            =   " "
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Vida Util"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label24 
         Caption         =   "Embalaje"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   555
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "N.Unidas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "F.Intervencion"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label45 
         Caption         =   "Estado"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   61
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label46 
         Caption         =   "Carga"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   60
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label57 
         Caption         =   "Riesgo"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label56 
         Caption         =   "R.Sec."
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1150
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Clase"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label44 
         Caption         =   "Caracteristicas"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.CommandButton Bloqueo 
      Caption         =   "Bloqueo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   44
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   43
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton AvisoError 
      Caption         =   "Sistema sin Conexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      Picture         =   "compo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Caratula 
      Caption         =   "Caratula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8400
      TabIndex        =   37
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Observaciones1 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   36
      Text            =   " "
      Top             =   7320
      Width           =   6135
   End
   Begin VB.TextBox Observaciones2 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   7680
      Width           =   6135
   End
   Begin VB.TextBox Referencia 
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   33
      Text            =   " "
      Top             =   6960
      Width           =   1215
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
      Height          =   1230
      Left            =   4440
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   2280
      TabIndex        =   15
      Top             =   720
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
      Height          =   450
      Left            =   120
      TabIndex        =   14
      Top             =   720
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
      Height          =   450
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   9495
      Begin VB.TextBox WTipo 
         Height          =   300
         Left            =   360
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "  "
         Top             =   480
         Width           =   495
      End
      Begin MSMask.MaskEdBox WArticulo2 
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   0
         TabIndex        =   13
         Text            =   " "
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSMask.MaskEdBox WArticulo1 
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox WCantidad 
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   9
         Text            =   " "
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Left            =   8160
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desripcion"
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
         TabIndex        =   21
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto Terminado"
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
         Left            =   2280
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materia Prima"
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
         Left            =   840
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label WDescripcion 
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
         Height          =   300
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   4335
      End
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   360
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
      Height          =   1230
      ItemData        =   "compo.frx":0742
      Left            =   4440
      List            =   "compo.frx":0749
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
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
      Height          =   420
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   975
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
      Height          =   420
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid WGrilla 
      Height          =   3855
      Left            =   120
      TabIndex        =   67
      Top             =   2040
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6800
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label12 
      Caption         =   "Control de Cambios"
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
      TabIndex        =   45
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Responsable"
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
      TabIndex        =   42
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label DesOperador 
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
      Left            =   6840
      TabIndex        =   41
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      TabIndex        =   40
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label ZEstado 
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
      Left            =   8760
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "OBservaciones"
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
      TabIndex        =   34
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Ref. Laboratorio"
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
      TabIndex        =   32
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Xversion 
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
      Left            =   1440
      TabIndex        =   27
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Version"
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
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label WFecha 
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
      Left            =   3720
      TabIndex        =   25
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha "
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
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label DesTerminado 
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCompo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vector(100, 10) As String
Dim ZVector(100, 15) As String
Dim CargaEmpresa(12, 2) As String
Private Auxiliar(120, 7) As String

Dim ZZImpreProceso As Integer

Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstComposicionVersion As Recordset
Dim spComposicionVerion As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstCaratula As Recordset
Dim spCaratula As String
Dim rstOperador As Recordset
Dim spOperador As String

Private Producto As String
Private Costo As Double


Dim XParam As String
Dim ZVersion As String
Dim ZRenglon As String
Dim ZOperador As String
Dim WProceso As Integer

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Auxi As String
Private Clave As String
Private WGraba As String
Private WVersion As Single
Private XVector(100, 20) As String

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim WNombreEmail As String
Dim MAttach As String

Dim ZZImpreAnterior  As String

Dim ZZRestriccion As Integer
Dim WRestriccion As String


Private Sub Bloqueo_Click()
    Rem On Error GoTo WError
    
    If Val(WEmpresa) <> 1 And Val(WEmpresa) <> 8 Then
        Exit Sub
    End If
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1
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
        Case 2
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
        Case 3
            CargaEmpresa(1, 1) = "0003"
            CargaEmpresa(1, 2) = "Empresa03"
        Case 4
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 5
            CargaEmpresa(1, 1) = "0005"
            CargaEmpresa(1, 2) = "Empresa05"
        Case 6
            CargaEmpresa(1, 1) = "0006"
            CargaEmpresa(1, 2) = "Empresa06"
        Case 7
            CargaEmpresa(1, 1) = "0007"
            CargaEmpresa(1, 2) = "Empresa07"
        Case 8
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 9
            CargaEmpresa(1, 1) = "0009"
            CargaEmpresa(1, 2) = "Empresa09"
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case 11
            CargaEmpresa(1, 1) = "0011"
            CargaEmpresa(1, 2) = "Empresa11"
        Case Else
    End Select
                
    For Cicla = 1 To 7
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub
    
    If WGraba <> "S" Then
    
        WProceso = 2
        Call Ingresa_clave

               Else

        Terminado.Text = UCase(Terminado.Text)
        
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(WEmpresa)
            Case 1
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
            Case 2
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
            Case 3
                CargaEmpresa(1, 1) = "0003"
                CargaEmpresa(1, 2) = "Empresa03"
            Case 4
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 5
                CargaEmpresa(1, 1) = "0005"
                CargaEmpresa(1, 2) = "Empresa05"
            Case 6
                CargaEmpresa(1, 1) = "0006"
                CargaEmpresa(1, 2) = "Empresa06"
            Case 7
                CargaEmpresa(1, 1) = "0007"
                CargaEmpresa(1, 2) = "Empresa07"
            Case 8
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 9
                CargaEmpresa(1, 1) = "0009"
                CargaEmpresa(1, 2) = "Empresa09"
            Case 10
                CargaEmpresa(1, 1) = "0010"
                CargaEmpresa(1, 2) = "Empresa10"
            Case 11
                CargaEmpresa(1, 1) = "0011"
                CargaEmpresa(1, 2) = "Empresa11"
            Case Else
        End Select
                
        For Cicla = 1 To 7
            If CargaEmpresa(Cicla, 1) <> "" Then
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " Estado = " + "'" + "N" + "',"
                ZSql = ZSql + " EstadoI = " + "'" + "N" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                            
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
            End If
        Next Cicla
        
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
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        Call Limpia_Click
        
    End If
    
    Exit Sub
    
Control_error:
    Beep
    
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next

End Sub

Private Sub Borra_Click()

    WGrilla.Col = 1
    WGrilla.Text = ""
    
    WGrilla.Col = 2
    WGrilla.Text = ""

    WGrilla.Col = 3
    WGrilla.Text = ""
    
    WGrilla.Col = 4
    WGrilla.Text = ""

    WGrilla.Col = 5
    WGrilla.Text = ""
    
    WTipo.SetFocus
End Sub

Private Sub Caratula_Click()
    ZZImpreProceso = 0
    Call ImpreCaratula
End Sub

Private Sub CaratulaPDF_Click()
    ZZImpreProceso = 1
    Call ImpreCaratula
End Sub

Private Sub ImpreCaratula()

    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
    
    CargaEmpresa(1, 1) = "0003"
    CargaEmpresa(1, 2) = "Empresa03"
    CargaEmpresa(2, 1) = "0004"
    CargaEmpresa(2, 2) = "Empresa04"
                
    For Cicla = 1 To 2
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub

    Sql1 = "DELETE Caratula"
    spCaratula = Sql1
    Set rstCaratula = db.OpenRecordset(spCaratula, dbOpenSnapshot, dbSQLPassThrough)

    Erase XVector
    
    Renglon = 0
    Renglon1 = 0
    Renglon2 = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Composicion"
    Sql3 = " Where Composicion.Terminado = " + "'" + Terminado.Text + "'"
    Sql4 = " Order by Composicion.Clave"
    spComposicion = Sql1 + Sql2 + Sql3 + Sql4
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    If rstComposicion.RecordCount > 0 Then
    With rstComposicion
    
        .MoveFirst
        If .NoMatch = False Then
        
            Do
                ZZEntraCompo = "S"
                
                If rstComposicion!Tipo = "M" Then
                    If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                        ZZEntraCompo = "N"
                    End If
                End If
                
                If ZZEntraCompo = "S" Then

                    Renglon1 = Renglon1 + 1
                        
                    XVector(Renglon1, 1) = rstComposicion!Tipo
                    XVector(Renglon1, 2) = rstComposicion!Articulo1
                    XVector(Renglon1, 3) = rstComposicion!Articulo2
                    XVector(Renglon1, 4) = Str$(rstComposicion!Cantidad)
                    XVector(Renglon1, 5) = rstComposicion!Clave
                    XVector(Renglon1, 6) = rstComposicion!Terminado
                    
                End If
                    
                .MoveNext
                    
                If .EOF = True Then
                    Exit Do
                End If
                        
            Loop
            
        End If
            
    End With
    rstComposicion.Close
    
    End If
    
    XXXCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If XXXCodigo >= 25000 And XXXCodigo <= 25999 Then
    
            Else
        
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
        
        Sql1 = "Select *"
        Sql2 = " FROM EspecifUnifica"
        Sql3 = " Where EspecifUnifica.Producto = " + "'" + Terminado.Text + "'"
        spEspecifUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecifUnifica.RecordCount > 0 Then
            If rstEspecifUnifica!Ensayo1 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo1
                XVector(Renglon2, 9) = rstEspecifUnifica!Valor1
                WValor = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo2 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo2
                XVector(Renglon2, 9) = rstEspecifUnifica!valor2
                WValor = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo3 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo3
                XVector(Renglon2, 9) = rstEspecifUnifica!Valor3
                WValor = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo4 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo4
                XVector(Renglon2, 9) = rstEspecifUnifica!valor4
                WValor = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo5 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo5
                XVector(Renglon2, 9) = rstEspecifUnifica!valor5
                WValor = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo6 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo6
                XVector(Renglon2, 9) = rstEspecifUnifica!valor6
                WValor = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo7 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo7
                XVector(Renglon2, 9) = rstEspecifUnifica!valor7
                WValor = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo8 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo8
                XVector(Renglon2, 9) = rstEspecifUnifica!valor8
                WValor = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo9 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo9
                XVector(Renglon2, 9) = rstEspecifUnifica!valor9
                WValor = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            If rstEspecifUnifica!Ensayo10 <> 0 Then
                Renglon2 = Renglon2 + 1
                XVector(Renglon2, 7) = rstEspecifUnifica!Ensayo10
                XVector(Renglon2, 9) = rstEspecifUnifica!valor10
                WValor = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                WValor = Trim(WValor)
                If WValor <> "" Then
                    Renglon2 = Renglon2 + 1
                    XVector(Renglon2, 7) = ""
                    XVector(Renglon2, 8) = ""
                    XVector(Renglon2, 9) = WValor
                End If
            End If
        
            rstEspecifUnifica.Close
        End If
        
        Call Conecta_Empresa
        
    End If
        
    If Renglon1 > Renglon2 Then
        WRenglon = Renglon1
            Else
        WRenglon = Renglon2
    End If
    
    For DA = 1 To WRenglon
        
        Tipo = XVector(DA, 1)
        Articulo1 = XVector(DA, 2)
        Articulo2 = XVector(DA, 3)
        Cantidad = Val(XVector(DA, 4))
        WClave = XVector(DA, 5)
        WTerminado = XVector(DA, 6)
        WEnsayo = XVector(DA, 7)
        WResultado = Trim(XVector(DA, 9))
        
        Renglon = Renglon + 1
        Auxi = Str$(Renglon)
        Call Ceros(Auxi, 2)
                    
        WTerminado = Terminado.Text
        WRenglon = Auxi
        WTipo = Tipo
        If Tipo = "M" Then
            XArticulo1 = Articulo1
            XArticulo2 = "  -     -   "
                Else
            XArticulo1 = "  -   -  "
            XArticulo2 = Articulo2
        End If
        WCantidad = Str$(Cantidad)
        WClave = WTerminado + WRenglon
        
        WDescriterminado = ""
        WDescriarticulo1 = ""
        WDescriarticulo2 = ""
        WDesensayo = ""
        
        Select Case Tipo
            Case "T"
                Producto = Articulo2
                spTerminado = "ConsultaTerminado " + "'" + Articulo2 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescriarticulo2 = Trim(Left$(rstTerminado!Descripcion, 30))
                    rstTerminado.Close
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriarticulo1 = Trim(Left$(rstArticulo!Descripcion, 30))
                    rstArticulo.Close
                End If
            Case Else
        End Select
        
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescriterminado = Trim(Left$(rstTerminado!Descripcion, 30))
            rstTerminado.Close
        End If
        
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
        
        spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            WDesensayo = Trim(Left$(rstEnsayo!Descripcion, 30))
            rstEnsayo.Close
        End If
        
        Call Conecta_Empresa
                        
        Sql1 = "INSERT INTO Caratula ("
        Sql2 = "Clave ,"
        Sql3 = "Terminado ,"
        Sql4 = "Renglon ,"
        Sql5 = "Tipo ,"
        Sql6 = "Articulo1 ,"
        Sql7 = "Articulo2 ,"
        Sql8 = "Cantidad ,"
        Sql9 = "DescriTerminado ,"
        Sql10 = "DescriArticulo1 ,"
        Sql11 = "DescriArticulo2 ,"
        Sql12 = "Referencia ,"
        Sql13 = "Observaciones1 ,"
        Sql14 = "Observaciones2 ,"
        Sql15 = "Ensayo ,"
        Sql16 = "DesEnsayo ,"
        Sql17 = "Resultado) "
        Sql18 = "Values ("
        Sql19 = "'" + WClave + "',"
        Sql20 = "'" + WTerminado + "',"
        Sql21 = "'" + WRenglon + "',"
        Sql22 = "'" + WTipo + "',"
        Sql23 = "'" + XArticulo1 + "',"
        Sql24 = "'" + XArticulo2 + "',"
        Sql25 = "'" + WCantidad + "',"
        Sql26 = "'" + Left$(WDescriterminado, 30) + "',"
        Sql27 = "'" + Left$(WDescriarticulo1, 30) + "',"
        Sql28 = "'" + Left$(WDescriarticulo2, 30) + "',"
        Sql29 = "'" + Left$(Referencia.Text, 10) + "',"
        Sql30 = "'" + Left$(Observaciones1.Text, 50) + "',"
        Sql31 = "'" + Left$(Observaciones2.Text, 50) + "',"
        Sql32 = "'" + Left$(WEnsayo, 50) + "',"
        Sql33 = "'" + Left$(WDesensayo, 50) + "',"
        Sql34 = "'" + Left$(WResultado, 50) + "')"
       
        spCaratula = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                     Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                     Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                     Sql31 + Sql32 + Sql33 + Sql34
        Set rstCaratula = db.OpenRecordset(spCaratula, dbOpenSnapshot, dbSQLPassThrough)
        
    Next DA
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WNombreEmpresa = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WNombreEmpresa
            .Update
        End If
    End With
    

    Listado.WindowTitle = "Caratula de Productos Terminados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CARATULA.terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.Destination = 1
   Rem  Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    XXXCodigo = Val(Mid$(Terminado.Text, 4, 5))
    If XXXCodigo >= 25000 And XXXCodigo <= 25999 Then
    
        Listado.SQLQuery = "SELECT Caratula.Terminado, Caratula.Renglon, Caratula.Tipo, Caratula.Articulo1, Caratula.Articulo2, Caratula.Cantidad, Caratula.DescriTerminado, Caratula.DescriArticulo1, Caratula.DescriArticulo2, Caratula.Referencia, Caratula.Observaciones1, Caratula.Observaciones2, Caratula.Ensayo, Caratula.DesEnsayo, Caratula.Resultado, " _
                        + "Terminado.Clase, Terminado.Intervencion, Terminado.Naciones, Terminado.Version, Terminado.FechaVersion, Terminado.Escrito, Terminado.Vida  " _
                        + "From " _
                        + DSQ + ".dbo.Caratula Caratula, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Caratula.Terminado = Terminado.Codigo AND " _
                        + "Caratula.Terminado >= '" + Terminado.Text + "' AND " _
                        + "Caratula.Terminado <= '" + Terminado.Text + "'"
        Listado.ReportFileName = "CaratulaFarma.rpt"
        
            Else
            
        Listado.SQLQuery = "SELECT Caratula.Terminado, Caratula.Renglon, Caratula.Tipo, Caratula.Articulo1, Caratula.Articulo2, Caratula.Cantidad, Caratula.DescriTerminado, Caratula.DescriArticulo1, Caratula.DescriArticulo2, Caratula.Referencia, Caratula.Observaciones1, Caratula.Observaciones2, Caratula.Ensayo, Caratula.DesEnsayo, Caratula.Resultado, " _
                        + "Terminado.Clase, Terminado.Intervencion, Terminado.Naciones, Terminado.Version, Terminado.FechaVersion, Terminado.Escrito, Terminado.Vida " _
                        + "From " _
                        + DSQ + ".dbo.Caratula Caratula, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Caratula.Terminado = Terminado.Codigo AND " _
                        + "Caratula.Terminado >= '" + Terminado.Text + "' AND " _
                        + "Caratula.Terminado <= '" + Terminado.Text + "'"
        Listado.ReportFileName = "Caratula.rpt"
        
    End If
    
    ZZImpreAnterior = Printer.DeviceName
    
    If ZZImpreProceso = 1 Then
        Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "CutePDF Writer" + Chr$(34)
    End If
    
   Rem Listado.Destination = 0
    Listado.Action = 1
    
    If ZZImpreProceso = 1 Then
        Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + ZZImpreAnterior + Chr$(34)
    End If
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
    
End Sub

Private Sub CierraPanta_Click()
    PantaExplosion.Visible = False
End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click
    PrgCompo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()

    Dim ZZVectorI(1000) As String
    Dim ZZVectorII(10000) As String
    Dim ZZRestriccion As Integer
    
    Erase ZZVectorI
    Erase ZZVectorII
    ZZLugarI = 0
    ZZLugarII = 0

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                ZZRestriccion = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                If ZZRestriccion = 1 Then
                    ZZLugarI = ZZLugarI + 1
                    ZZVectorI(ZZLugarI) = rstArticulo!Codigo
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
    
    For Ciclo = 1 To ZZLugarI
    
        ZZArticulo = ZZVectorI(Ciclo)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Composicion"
        ZSql = ZSql + " Where Composicion.Articulo1 = " + "'" + ZZArticulo + "'"
        spComposicion = ZSql
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "S"
                        
                        For CicloII = 1 To ZZLugarII
                            If ZZVectorII(CicloII) = rstComposicion!Terminado Then
                                Entra = "N"
                                Exit For
                            End If
                        Next CicloII
                        
                        If Entra = "S" Then
                            ZZLugarII = ZZLugarII + 1
                            ZZVectorII(ZZLugarII) = rstComposicion!Terminado
                        End If
                            
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
        End If
        
    Next Ciclo
    
    
    XEmpresa = WEmpresa
    
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
                
    For Cicla = 1 To 7
        If CargaEmpresa(Cicla, 1) <> "" Then
        
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            For Ciclo = 1 To ZZLugarII
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " Restriccion = " + "'" + "1" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + ZZVectorII(Ciclo) + "'"
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
            
        End If
    Next Cicla
    
    Call Conecta_Empresa

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados (Cabecera)"
     Opcion.AddItem "Materia Prima"
     Opcion.AddItem "Productos Terminados (Cuerpo)"

     Opcion.Visible = True
     
 End Sub


Private Sub Explosion_Click()

    Dim ZZTerminado(1000, 2) As String
    Dim ZZArticulo(1000, 2) As String
    Dim ZZLugarTermi As Integer
    Dim ZZLugarArti As Integer
    
    Erase ZZTerminado
    Erase ZZArticulo
    
    ZZLugarTermi = 0
    ZZLugarArti = 0
    
    ZZLugarTermi = 1
    ZZTerminado(ZZLugarTermi, 1) = Terminado.Text
    ZZTerminado(ZZLugarTermi, 2) = "1"
    Ciclo = 1
    
    Do
    
        WTerminado = ZZTerminado(Ciclo, 1)
        WWCantidad = Val(ZZTerminado(Ciclo, 2))
        
        spComposicion = "ConsultaComposicionProducto " + "'" + WTerminado + "'"
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstComposicion.RecordCount > 0 Then
        
            With rstComposicion
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                        If rstComposicion!Tipo = "M" Then
                        
                            ZZEntra = "S"
                            For CicloII = 1 To ZZLugarArti
                                If ZZArticulo(CicloII, 1) = rstComposicion!Articulo1 Then
                                    ZZArticulo(CicloII, 2) = Str$(Val(ZZArticulo(CicloII, 2)) + (rstComposicion!Cantidad * WWCantidad))
                                    ZZEntra = "N"
                                    Exit For
                                End If
                            Next CicloII
                            If ZZEntra = "S" Then
                                ZZLugarArti = ZZLugarArti + 1
                                ZZArticulo(ZZLugarArti, 1) = rstComposicion!Articulo1
                                ZZArticulo(ZZLugarArti, 2) = Str$(rstComposicion!Cantidad * WWCantidad)
                            End If
                                Else
                            ZZLugarTermi = ZZLugarTermi + 1
                            ZZTerminado(ZZLugarTermi, 1) = rstComposicion!Articulo2
                            ZZTerminado(ZZLugarTermi, 2) = Str$(rstComposicion!Cantidad * WWCantidad)
                        End If
                        
                        .MoveNext
                        
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                End If
                
            End With
            rstComposicion.Close
    
        End If
        
        Ciclo = Ciclo + 1
        
        If Ciclo > ZZLugarTermi Then
            Exit Do
        End If
        
    Loop
    
    Call Limpia_VectorII
    
    ZSuma = 0
    
    For Ciclo = 1 To ZZLugarArti
    
        WGrillaExplo.TextMatrix(Ciclo, 1) = ZZArticulo(Ciclo, 1)
        WGrillaExplo.TextMatrix(Ciclo, 3) = ZZArticulo(Ciclo, 2)
        
        WGrillaExplo.TextMatrix(Ciclo, 3) = Pusing("###,###.#####", WGrillaExplo.TextMatrix(Ciclo, 3))
        
        ZSuma = ZSuma + Val(WGrillaExplo.TextMatrix(Ciclo, 3))
        
        spArticulo = "ConsultaArticulo " + "'" + ZZArticulo(Ciclo, 1) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WGrillaExplo.TextMatrix(Ciclo, 2) = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    
    Next Ciclo
    
    PantaExplosion.Height = 5295
    PantaExplosion.Left = 360
    PantaExplosion.Top = 1200
    PantaExplosion.Width = 9615
    
    PantaExplosion.Visible = True
    
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If


    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 8 Then
        Graba.Enabled = True
            Else
        Graba.Enabled = False
    End If
End Sub

Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0, 2
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            
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
        
        Case 1
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub WGrilla_GotFocus()

    WGrilla.Col = 1
    WTipo.Text = WGrilla.Text
    
    If Len(WTipo.Text) <> 0 Then
        WLinea.Text = WGrilla.Row
            Else
        WLinea.Text = ""
    End If
    
    WGrilla.Col = 2
    If Len(WGrilla.Text) = 10 Then
        WArticulo1.Text = WGrilla.Text
            Else
        WArticulo1.Text = "  -   -   "
    End If
    
    WGrilla.Col = 3
    If Len(WGrilla.Text) = 12 Then
        WArticulo2.Text = WGrilla.Text
            Else
        WArticulo2.Text = "  -     -   "
    End If
    
    WGrilla.Col = 4
    WDescripcion.Caption = WGrilla.Text

    WGrilla.Col = 5
    Call Conver(WGrilla.Text, Auxi)
    If Val(Auxi) = 0 Then
        WCantidad.Text = ""
            Else
        WCantidad.Text = Auxi
    End If

    WTipo.SetFocus

End Sub

Private Sub Graba_Click()

    Rem On Error GoTo WError
    
    If Trim(Observaciones2.Text) = "" Then
        m$ = "Se debe informar el campo Control de CAmbio"
        a% = MsgBox(m$, 0, "Composicion de Producto Terminado")
        Exit Sub
    End If
    
    If Val(WEmpresa) <> 1 And Val(WEmpresa) <> 8 Then
        Exit Sub
    End If
    
    ZZTrabaTerminado = 0
    
    For iRow = 1 To 50
    
        Tipo = UCase(WGrilla.TextMatrix(iRow, 1))
        Articulo1 = UCase(WGrilla.TextMatrix(iRow, 2))
        Articulo2 = UCase(WGrilla.TextMatrix(iRow, 3))
        Auxi1 = WGrilla.TextMatrix(iRow, 5)
        Call Conver(Auxi1, Auxi)
        Cantidad = Val(Auxi)
            
        If Cantidad <> 0 Then
            If Tipo = "M" Then
            
                If Left$(Articulo1, 2) = "DK" Then
                    Exit Sub
                End If
                spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZRestriccion = IIf(IsNull(rstArticulo!Restriccion), "0", rstArticulo!Restriccion)
                    rstArticulo.Close
                    If ZZRestriccion = 1 Then
                        ZZTrabaTerminado = 1
                    End If
                        Else
                    Exit Sub
                End If
            
                    Else
                    
                If Left$(Articulo2, 2) = "NK" Or Left$(Articulo2, 2) = "RE" Then
                    Exit Sub
                End If
                spTerminado = "ConsultaTerminado " + "'" + Articulo2 + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZZRestriccion = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
                    rstTerminado.Close
                    If ZZRestriccion = 1 Then
                        ZZTrabaTerminado = 1
                    End If
                        Else
                    Exit Sub
                End If
            
            End If
                
        End If
                                
    Next iRow
    
    If ZZTrabaTerminado = 1 Then
        T$ = "Composicion de Producto terminmado"
        m$ = "Esta formula contiene sustancias restringidas. Desea grabarlas?"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
                Else
            Exit Sub
        End If
    End If
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1
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
        Case 2
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
        Case 3
            CargaEmpresa(1, 1) = "0003"
            CargaEmpresa(1, 2) = "Empresa03"
        Case 4
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 5
            CargaEmpresa(1, 1) = "0005"
            CargaEmpresa(1, 2) = "Empresa05"
        Case 6
            CargaEmpresa(1, 1) = "0006"
            CargaEmpresa(1, 2) = "Empresa06"
        Case 7
            CargaEmpresa(1, 1) = "0007"
            CargaEmpresa(1, 2) = "Empresa07"
        Case 8
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 9
            CargaEmpresa(1, 1) = "0009"
            CargaEmpresa(1, 2) = "Empresa09"
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case 11
            CargaEmpresa(1, 1) = "0011"
            CargaEmpresa(1, 2) = "Empresa11"
        Case Else
    End Select
                
    For Cicla = 1 To 7
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub
    
    If WGraba <> "S" Then
    
        WProceso = 0
        Call Ingresa_clave

               Else

        Terminado.Text = UCase(Terminado.Text)
        
        Producto = Terminado.Text
        Call Calcula_Costo(Producto, Costo)
        CostoAnterior = Costo
        
        Erase ZVector
        ZLugar = 0
        
        spComposicion = "ConsultaComposicionProducto " + "'" + Terminado.Text + "'"
        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar, 1) = rstComposicion!Terminado
                        ZVector(ZLugar, 2) = rstComposicion!Renglon
                        ZVector(ZLugar, 3) = rstComposicion!Tipo
                        ZVector(ZLugar, 4) = rstComposicion!Articulo1
                        ZVector(ZLugar, 5) = rstComposicion!Articulo2
                        ZVector(ZLugar, 6) = Str(rstComposicion!Cantidad)
                        ZVector(ZLugar, 7) = IIf(IsNull(rstComposicion!Referencia), "", rstComposicion!Referencia)
                        ZVector(ZLugar, 8) = IIf(IsNull(rstComposicion!Observaciones1), "", rstComposicion!Observaciones1)
                        ZVector(ZLugar, 9) = IIf(IsNull(rstComposicion!Observaciones2), "", rstComposicion!Observaciones2)
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstComposicion.Close
        End If
        
        If ZLugar > 0 Then
        
            spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                ZVersion = "0"
                ZFechaInicio = ""
                ZFechaFinal = ""
                ZVersion = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
                ZFechaInicio = IIf(IsNull(rstTerminado!FechaVersion), "", rstTerminado!FechaVersion)
                ZFechaInicio = Mid$(ZFechaInicio, 4, 2) + "/" + Left$(ZFechaInicio, 2) + "/" + Right$(ZFechaInicio, 4)
                ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                rstTerminado.Close
            End If
            
            For Ciclo = 1 To ZLugar
        
                ZTerminado = ZVector(Ciclo, 1)
                ZRenglon = ZVector(Ciclo, 2)
                ZTipo = ZVector(Ciclo, 3)
                ZArticulo1 = ZVector(Ciclo, 4)
                ZArticulo2 = ZVector(Ciclo, 5)
                ZCantidad = ZVector(Ciclo, 6)
                ZReferencia = ZVector(Ciclo, 7)
                ZObservaciones1 = ZVector(Ciclo, 8)
                ZObservaciones2 = ZVector(Ciclo, 9)
            
                Call Ceros(ZVersion, 4)
                Call Ceros(ZRenglon, 2)
                ZClave = ZTerminado + ZVersion + ZRenglon
                
                ZSql = ""
                ZSql = ZSql & "INSERT INTO ComposicionVersion ("
                ZSql = ZSql & "Clave, "
                ZSql = ZSql & "Terminado, "
                ZSql = ZSql & "Version, "
                ZSql = ZSql & "Renglon, "
                ZSql = ZSql & "Tipo, "
                ZSql = ZSql & "Articulo1 , "
                ZSql = ZSql & "Articulo2 , "
                ZSql = ZSql & "Cantidad , "
                ZSql = ZSql & "FechaInicio , "
                ZSql = ZSql & "FechaFinal , "
                ZSql = ZSql & "Referencia , "
                ZSql = ZSql & "Observaciones1 , "
                ZSql = ZSql & "Observaciones2) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + ZClave + "',"
                ZSql = ZSql & "'" + ZTerminado + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZRenglon + "',"
                ZSql = ZSql & "'" + ZTipo + "',"
                ZSql = ZSql & "'" + ZArticulo1 + "',"
                ZSql = ZSql & "'" + ZArticulo2 + "',"
                ZSql = ZSql & "'" + ZCantidad + "',"
                ZSql = ZSql & "'" + ZFechaInicio + "',"
                ZSql = ZSql & "'" + ZFechaFinal + "',"
                ZSql = ZSql & "'" + ZReferencia + "',"
                ZSql = ZSql & "'" + ZObservaciones1 + "',"
                ZSql = ZSql & "'" + ZObservaciones2 + "')"
          
                spComposicionVersion = ZSql
                Set rstComposicionVersion = db.OpenRecordset(spComposicionVersion, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
            
        End If
            
        Renglon = 0
        Erase Vector
        
        For iRow = 1 To 50
        
            WRow = iRow
            WGrilla.Row = WRow
                
            WGrilla.Col = 1
            Tipo = WGrilla.Text
                
            WGrilla.Col = 2
            Articulo1 = UCase(WGrilla.Text)
                
            WGrilla.Col = 3
            Articulo2 = UCase(WGrilla.Text)
                
            WGrilla.Col = 5
            Call Conver(WGrilla.Text, Auxi)
            Cantidad = Val(Auxi)
                
            If Cantidad <> 0 Then
                Renglon = Renglon + 1
                Vector(Renglon, 1) = Tipo
                Vector(Renglon, 2) = Articulo1
                Vector(Renglon, 3) = Articulo2
                Vector(Renglon, 4) = Str$(Cantidad)
            End If
                                    
        Next iRow
        
        Counter = Renglon
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(WEmpresa)
            Case 1
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
            Case 2
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
            Case 3
                CargaEmpresa(1, 1) = "0003"
                CargaEmpresa(1, 2) = "Empresa03"
            Case 4
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 5
                CargaEmpresa(1, 1) = "0005"
                CargaEmpresa(1, 2) = "Empresa05"
            Case 6
                CargaEmpresa(1, 1) = "0006"
                CargaEmpresa(1, 2) = "Empresa06"
            Case 7
                CargaEmpresa(1, 1) = "0007"
                CargaEmpresa(1, 2) = "Empresa07"
            Case 8
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 9
                CargaEmpresa(1, 1) = "0009"
                CargaEmpresa(1, 2) = "Empresa09"
            Case 10
                CargaEmpresa(1, 1) = "0010"
                CargaEmpresa(1, 2) = "Empresa10"
            Case 11
                CargaEmpresa(1, 1) = "0011"
                CargaEmpresa(1, 2) = "Empresa11"
            Case Else
        End Select
                
        For Cicla = 1 To 7
            If CargaEmpresa(Cicla, 1) <> "" Then
            
                Renglon = 0
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                spComposicion = "BorrarComposicion " + "'" + Terminado.Text + "'"
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                
                For Lugar = 1 To Counter
        
                    Tipo = Vector(Lugar, 1)
                    Articulo1 = Vector(Lugar, 2)
                    Articulo2 = Vector(Lugar, 3)
                    Cantidad = Val(Vector(Lugar, 4))
                    
                    If Cantidad <> 0 Then
                    
                        Renglon = Renglon + 1
                        Auxi = Str$(Renglon)
                        Call Ceros(Auxi, 2)
                    
                        WTerminado = Terminado.Text
                        WRenglon = Auxi
                        WTipo = Tipo
                        If Tipo = "M" Then
                            XArticulo1 = Articulo1
                            XArticulo2 = "  -     -   "
                                Else
                            XArticulo1 = "  -   -  "
                            XArticulo2 = Articulo2
                        End If
                        WCantidad = Str$(Cantidad)
                        WClave = WTerminado + WRenglon
                        XDate = Date$
                        WCosto1 = "0"
                        WCosto2 = "0"
                        
                        Sql1 = "INSERT INTO Composicion ("
                        Sql2 = "Clave ,"
                        Sql3 = "Terminado ,"
                        Sql4 = "Renglon ,"
                        Sql5 = "Tipo ,"
                        Sql6 = "Articulo1 ,"
                        Sql7 = "Articulo2 ,"
                        Sql8 = "Cantidad ,"
                        Sql9 = "WDate ,"
                        Sql10 = "Costo1 ,"
                        Sql11 = "Costo2 ,"
                        Sql12 = "Referencia ,"
                        Sql13 = "Observaciones1 ,"
                        Sql14 = "Observaciones2 )"
                        Sql15 = "Values ("
                        Sql16 = "'" + WClave + "',"
                        Sql17 = "'" + WTerminado + "',"
                        Sql18 = "'" + WRenglon + "',"
                        Sql19 = "'" + WTipo + "',"
                        Sql20 = "'" + XArticulo1 + "',"
                        Sql21 = "'" + XArticulo2 + "',"
                        Sql22 = "'" + WCantidad + "',"
                        Sql23 = "'" + XDate + "',"
                        Sql24 = "'" + WCosto1 + "',"
                        Sql25 = "'" + WCosto2 + "',"
                        Sql26 = "'" + Referencia.Text + "',"
                        Sql27 = "'" + Observaciones1.Text + "',"
                        Sql28 = "'" + Observaciones2.Text + "')"
        
                        spComposicion = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                                Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                                Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28
                        Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                                        
                Next Lugar
                
                
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Composicion SET "
                ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
                ZSql = ZSql + " Where Terminado = " + "'" + WTerminado + "'"
                            
                spComposicion = ZSql
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
                spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    
                    WVersion = 0
                    WVersion = IIf(IsNull(rstTerminado!Version), "0", rstTerminado!Version)
                    WVersion = WVersion + 1
                    XTerminado = Terminado.Text
                    Xversion = Str$(WVersion)
                    XFechaVersion = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    ZObserva = ""
                    rstTerminado.Close
                            
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Terminado SET "
                    ZSql = ZSql + " Restriccion = " + "'" + Str$(ZZTrabaTerminado) + "',"
                    ZSql = ZSql + " Version = " + "'" + Xversion + "',"
                    ZSql = ZSql + " FechaVersion = " + "'" + XFechaVersion + "',"
                    ZSql = ZSql + " Estado = " + "'" + "S" + "',"
                    ZSql = ZSql + " EstadoI = " + "'" + "N" + "',"
                    ZSql = ZSql + " EstadoII = " + "'" + "N" + "',"
                    ZSql = ZSql + " Observa = " + "'" + ZObserva + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + XTerminado + "'"
                            
                    spTerminado = ZSql
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
        
            End If
        Next Cicla
        
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
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        Call CaratulaPDF_Click
        
        Producto = Terminado.Text
        Call Calcula_Costo(Producto, Costo)
        CostoActual = Costo
        
        Rem If CostoAnterior <> CostoActual Then
        
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            
                ImpreCosto1 = Pusing("###,###.#####", Str$(CostoAnterior))
                ImpreCosto2 = Pusing("###,###.#####", Str$(CostoActual))
                XFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
                sTo = "lsantos@surfactan.com.ar;dsuarez@surfactan.com.ar"
                sCC = ""
                sBCC = ""
                sSubject = "CAMBIO DE FORMULA DEL " + Terminado.Text
            
                sBody = "Fecha:" + XFecha + " - " + _
                        "Costo Anterior : " + ImpreCosto1 + " - " + _
                        "Costo Actual : " + ImpreCosto2
                    
                Rem  ret = Shell("Start.exe " _
                rem         & "mailto:" & """" & sTo & """" _
                rem        & "?Subject=" & """" & sSubject & """" _
                rem        & "&cc=" & """" & sCC & """" _
                rem        & "&bcc=" & """" & sBCC & """" _
                rem        & "&Body=" & """" & sBody & """" _
                rem     & "&File=" & """" & "c:\autoexec.bat" & """" _
                        , 0)
                 Rem by nan 23-4-2013
                EmailAddress = sTo
                CopiaAddress = sCC
                MSubject = sSubject
                MBody = sBody
                MAttach = ""
                MAttachI = ""
                MAttachII = ""
                MAttachIII = ""
                MAttachIV = ""
                MAttachVI = ""
                MAttachVII = ""
                MAttachVIII = ""
                
                SendEmail
                        
                        
                 Rem by nan
                        Else
                        
                ImpreCosto1 = Pusing("###,###.#####", Str$(CostoAnterior))
                ImpreCosto2 = Pusing("###,###.#####", Str$(CostoActual))
                XFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                
                sTo = "hgutierrez@pellital.com.ar"
                sCC = ""
                sBCC = ""
                sSubject = "CAMBIO DE FORMULA DEL " + Terminado.Text
                sBody = "Fecha:" + XFecha + " - " + _
                        "Costo Anterior : " + ImpreCosto1 + " - " + _
                        "Costo Actual : " + ImpreCosto2
                SFile = ""
        
                EmailAddress = sTo
                CopiaAddress = sCC
                MSubject = sSubject
                MBody = sBody
                MAttach = ""
                MAttachI = ""
                MAttachII = ""
                MAttachIII = ""
                MAttachIV = ""
                MAttachVI = ""
                MAttachVII = ""
                MAttachVIII = ""
                
                SendEmail
                
            End If
                        
        Rem End If
        
        Call Limpia_Click
        
    End If
    
    Exit Sub
    
Control_error:
    Beep
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next
        
End Sub

Private Sub Ingresa_Click()
    WLinea.Text = ""
    WTipo.Text = ""
    WArticulo1.Text = "  -   -   "
    WArticulo2.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    WTipo.SetFocus
End Sub

Private Sub Limpia_Click()

    AvisoError.Visible = False

    WLinea.Text = ""
    WTipo.Text = ""
    WArticulo1.Text = "  -   -   "
    WArticulo2.Text = "  -     -   "
    WDescripcion.Caption = ""
    WCantidad.Text = ""
    Referencia.Text = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Referencia.Text = ""
    WGraba = ""
    
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Xversion.Caption = ""
    WFecha.Caption = ""
    ZEstado.Caption = ""
    DesOperador.Caption = ""
    Restriccion.Value = 0
    
    Call Limpia_Vector
    
    Renglon = 0
    
    Carga.ListIndex = 0
    EstadoProducto.ListIndex = 0

    Terminado.SetFocus

End Sub

Private Sub Revalida_Click()

    Rem On Error GoTo WError
    
    If Val(WEmpresa) <> 1 And Val(WEmpresa) <> 8 Then
        Exit Sub
    End If
    
    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1
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
        Case 2
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
        Case 3
            CargaEmpresa(1, 1) = "0003"
            CargaEmpresa(1, 2) = "Empresa03"
        Case 4
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 5
            CargaEmpresa(1, 1) = "0005"
            CargaEmpresa(1, 2) = "Empresa05"
        Case 6
            CargaEmpresa(1, 1) = "0006"
            CargaEmpresa(1, 2) = "Empresa06"
        Case 7
            CargaEmpresa(1, 1) = "0007"
            CargaEmpresa(1, 2) = "Empresa07"
        Case 8
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
        Case 9
            CargaEmpresa(1, 1) = "0009"
            CargaEmpresa(1, 2) = "Empresa09"
        Case 10
            CargaEmpresa(1, 1) = "0010"
            CargaEmpresa(1, 2) = "Empresa10"
        Case 11
            CargaEmpresa(1, 1) = "0011"
            CargaEmpresa(1, 2) = "Empresa11"
        Case Else
    End Select
                
    For Cicla = 1 To 7
        If CargaEmpresa(Cicla, 1) <> "" Then
            WEmpresa = CargaEmpresa(Cicla, 1)
            txtOdbc = CargaEmpresa(Cicla, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    On Error GoTo 0
    If WSalidaError = "N" Then Exit Sub
    
    If WGraba <> "S" Then
    
        WProceso = 1
        Call Ingresa_clave

               Else

        Terminado.Text = UCase(Terminado.Text)
        
        XEmpresa = WEmpresa
        Erase CargaEmpresa
        
        Select Case Val(WEmpresa)
            Case 1
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
            Case 2
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
            Case 3
                CargaEmpresa(1, 1) = "0003"
                CargaEmpresa(1, 2) = "Empresa03"
            Case 4
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 5
                CargaEmpresa(1, 1) = "0005"
                CargaEmpresa(1, 2) = "Empresa05"
            Case 6
                CargaEmpresa(1, 1) = "0006"
                CargaEmpresa(1, 2) = "Empresa06"
            Case 7
                CargaEmpresa(1, 1) = "0007"
                CargaEmpresa(1, 2) = "Empresa07"
            Case 8
                CargaEmpresa(1, 1) = "0002"
                CargaEmpresa(1, 2) = "Empresa02"
                CargaEmpresa(2, 1) = "0004"
                CargaEmpresa(2, 2) = "Empresa04"
                CargaEmpresa(3, 1) = "0008"
                CargaEmpresa(3, 2) = "Empresa08"
                CargaEmpresa(4, 1) = "0009"
                CargaEmpresa(4, 2) = "Empresa09"
            Case 9
                CargaEmpresa(1, 1) = "0009"
                CargaEmpresa(1, 2) = "Empresa09"
            Case 10
                CargaEmpresa(1, 1) = "0010"
                CargaEmpresa(1, 2) = "Empresa10"
            Case 11
                CargaEmpresa(1, 1) = "0011"
                CargaEmpresa(1, 2) = "Empresa11"
            Case Else
        End Select
                
        For Cicla = 1 To 7
            If CargaEmpresa(Cicla, 1) <> "" Then
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Terminado SET "
                ZSql = ZSql + " Estado = " + "'" + "S" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Terminado.Text + "'"
                            
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
            End If
        Next Cicla
        
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
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
        
        Call Limpia_Click
        
    End If
    
    Exit Sub
    
Control_error:
    Beep
    Rem MsgBox Err.Description
    WSalidaError = "N"
    AvisoError.Visible = True
    Resume Next

End Sub

Private Sub WTipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If WTipo.Text = "M" Or WTipo.Text = "T" Then
            If WTipo.Text = "M" Then
                WArticulo2.Text = "  -     -   "
                WArticulo1.SetFocus
                    Else
                WArticulo1.Text = "  -   -   "
                WArticulo2.SetFocus
            End If
                Else
            WTipo.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo1.Text = UCase(WArticulo1.Text)
        WArticulo = WArticulo1.Text
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WDescripcion.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            WCantidad.SetFocus
                Else
            WArticulo1.SetFocus
        End If
    End If
End Sub

Private Sub WArticulo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo2.Text = UCase(WArticulo2.Text)
        WTerminado = WArticulo2.Text
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion.Caption = rstTerminado!Descripcion
            rstTerminado.Close
            WCantidad.SetFocus
                Else
            WArticulo2.SetFocus
        End If
    End If
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.#####", WCantidad.Text)
        Call Alta_Vector
        Call Ingresa_Click
        WTipo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
        
            Indice = Pantalla.ListIndex
            WTerminado = WIndice.List(Indice)
            Terminado.Text = WTerminado
            Call Terminado_KeyPress(13)
            Rem spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Rem Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstTerminado.RecordCount > 0 Then
            Rem     DesTerminado.Caption = rstTerminado!Descripcion
             Rem   Xversion.Caption = ""
            Rem     WFecha.Caption = ""
            Rem     Xversion.Caption = rstTerminado!Version
            Rem     WFecha.Caption = rstTerminado!fechaversion
            Rem     rstTerminado.Close
            Rem     Call Proceso_Click
                Rem wgrilla.FirstRow = 0
                Rem wgrilla.Col = 0
                Rem wgrilla.Row = 0
            Rem End If
            Rem WLinea.Text = ""
            Rem WTipo.Text = ""
            Rem WArticulo1.Text = "  -   -   "
            Rem WArticulo2.Text = "  -     -   "
            Rem WDescripcion.Caption = ""
            Rem WCantidad.Text = ""
            Rem WTipo.SetFocus
    
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo1.Text = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo1.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WTipo.Text = "M"
                WArticulo1.Text = rstArticulo!Codigo
                WDescripcion.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            Call Alta_Vector
            
        Case 2
            Indice = Pantalla.ListIndex
            WArticulo2.Text = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + WArticulo2.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WTipo.Text = "T"
                WArticulo2.Text = rstTerminado!Codigo
                WDescripcion.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            Call Alta_Vector
            Rem WCantidad.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Carga.Clear
    
    Carga.AddItem ""
    Carga.AddItem "Alcalino"
    Carga.AddItem "Acido"
    Carga.AddItem "No ionico/Alcalino"
    Carga.AddItem "Neutro"
    
    Carga.ListIndex = 0
    
    EstadoProducto.Clear
    
    EstadoProducto.AddItem ""
    EstadoProducto.AddItem "Polvo"
    EstadoProducto.AddItem "Liquido"
    EstadoProducto.AddItem "Metal"
    EstadoProducto.AddItem "Pasta"
    
    EstadoProducto.ListIndex = 0
    
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Xversion.Caption = ""
    WFecha.Caption = ""
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Referencia.Text = ""
    DesOperador.Caption = ""
    Restriccion.Value = 0
    
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 8 Then
        Graba.Enabled = True
            Else
        Graba.Enabled = False
    End If
    
End Sub

Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
    Call Limpia_Vector
    
    Renglon = 0
    Erase Vector
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Terminado.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstComposicion.RecordCount > 0 Then
    
    With rstComposicion
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Renglon = Renglon + 1
            
                Vector(Renglon, 1) = rstComposicion!Tipo
                Vector(Renglon, 2) = rstComposicion!Articulo1
                Vector(Renglon, 3) = rstComposicion!Articulo2
                Vector(Renglon, 4) = ""
                Vector(Renglon, 5) = rstComposicion!Cantidad
                
                Referencia.Text = IIf(IsNull(rstComposicion!Referencia), "", rstComposicion!Referencia)
                Observaciones1.Text = IIf(IsNull(rstComposicion!Observaciones1), "", rstComposicion!Observaciones1)
                Observaciones2.Text = IIf(IsNull(rstComposicion!Observaciones2), "", rstComposicion!Observaciones2)
                
                Referencia.Text = RTrim(Referencia.Text)
                Observaciones1.Text = RTrim(Observaciones1.Text)
                Observaciones2.Text = RTrim(Observaciones2.Text)
                ZOperador = IIf(IsNull(rstComposicion!Operador), "O", rstComposicion!Operador)
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    rstComposicion.Close
    
    End If
    
    Renglon = 0
    
    For XX = 1 To 100
    
        If Vector(XX, 5) <> "" Then
            
                Renglon = Renglon + 1
            
                WGrilla.Row = Renglon
                
                WGrilla.Col = 1
                WGrilla.Text = Vector(XX, 1)
                
                WGrilla.Col = 2
                WGrilla.Text = Vector(XX, 2)
                Articulo1 = Vector(XX, 2)
                
                WGrilla.Col = 3
                WGrilla.Text = Vector(XX, 3)
                Articulo2 = Vector(XX, 3)
                
                WGrilla.Col = 5
                WGrilla.Text = Pusing("###,###.#####", Vector(XX, 5))
                
                If Vector(XX, 1) = "M" Then
                
                    spArticulo = "ConsultaArticulo " + "'" + Articulo1 + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WGrilla.Col = 4
                        WGrilla.Text = rstArticulo!Descripcion
                        rstArticulo.Close
                    End If
                        
                        Else
                        
                    WTerminado = Articulo2
                    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WGrilla.Col = 4
                        WGrilla.Text = rstTerminado!Descripcion
                        rstTerminado.Close
                    End If
                    
                End If
        End If
        
    Next XX
    
    If Val(ZOperador) <> 0 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Operador = " + "'" + ZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            DesOperador.Caption = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
            rstOperador.Close
        End If
        
    End If

    WGrilla.TopRow = 1
    WGrilla.Col = 1
    WGrilla.Row = 1
    
    Terminado.SetFocus
    Exit Sub

WError:
    Resume Next

End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            WGrilla.Row = Renglon
                
            WGrilla.Col = 1
            WGrilla.Text = WTipo.Text
            
            WGrilla.Col = 2
            WGrilla.Text = WArticulo1.Text
            
            WGrilla.Col = 3
            WGrilla.Text = WArticulo2.Text
            
            WGrilla.Col = 4
            WGrilla.Text = WDescripcion.Caption
                
            WGrilla.Col = 5
            WGrilla.Text = Pusing("###,###.#####", WCantidad.Text)
            
            WLinea.Text = WGrilla.Row
            
            ZZTop = WGrilla.Row - 13
            If ZZTop > 0 Then
                WGrilla.TopRow = ZZTop
            End If
            
                Else
                
            WGrilla.Row = Val(WLinea.Text)
                
            WGrilla.Col = 1
            WGrilla.Text = WTipo.Text
            
            WGrilla.Col = 2
            WGrilla.Text = WArticulo1.Text
            
            WGrilla.Col = 3
            WGrilla.Text = WArticulo2.Text
            
            WGrilla.Col = 4
            WGrilla.Text = WDescripcion.Caption
                
            WGrilla.Col = 5
            WGrilla.Text = Pusing("###,###.#####", WCantidad.Text)
            
    End If
    

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)

    Rem On Error GoTo WError

    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        Terminado.Text = WTerminado
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            Xversion.Caption = ""
            WFecha.Caption = ""
            Xversion.Caption = IIf(IsNull(rstTerminado!Version), "", rstTerminado!Version)
            WFecha.Caption = IIf(IsNull(rstTerminado!FechaVersion), "", rstTerminado!FechaVersion)
            ZEstado.Caption = IIf(IsNull(rstTerminado!Estado), "", rstTerminado!Estado)
            
            Clase.Text = ""
            Secundario.Text = ""
            Riesgo.Text = ""
            Intervencion.Text = ""
            Naciones.Text = ""
            Embalaje.Text = ""
            Clase.Text = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
            Secundario.Text = IIf(IsNull(rstTerminado!Secundario), "", rstTerminado!Secundario)
            Riesgo.Text = IIf(IsNull(rstTerminado!Riesgo), "", rstTerminado!Riesgo)
            Intervencion.Text = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
            Naciones.Text = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
            Embalaje.Text = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
            Vida.Text = IIf(IsNull(rstTerminado!Vida), "", rstTerminado!Vida)
            
            Caracteristicas.Text = IIf(IsNull(rstTerminado!DescriOnu), "", rstTerminado!DescriOnu)
            Carga.ListIndex = IIf(IsNull(rstTerminado!Carga), "0", rstTerminado!Carga)
            EstadoProducto.ListIndex = IIf(IsNull(rstTerminado!EstadoProducto), "0", rstTerminado!EstadoProducto)
        
            Naciones.Text = Trim(Naciones.Text)
        
            ZZRestriccion = IIf(IsNull(rstTerminado!Restriccion), "0", rstTerminado!Restriccion)
            Restriccion.Value = ZZRestriccion
            
            rstTerminado.Close
            Call Proceso_Click
            Call Ingresa_Click
            WTipo.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Referencia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones1.SetFocus
    End If
End Sub

Private Sub Observaciones1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones2.SetFocus
    End If
End Sub

Private Sub Observaciones2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Referencia.SetFocus
    End If
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    XClave.Visible = False

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        WGrabai = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            WGrabai = IIf(IsNull(rstOperador!GrabaI), "", rstOperador!GrabaI)
            rstOperador.Close
        End If
        
        If WGrabai = "S" Then
            WGraba = "S"
            XClave.Visible = False
            If WProceso = 0 Then
                Call Graba_Click
                    Else
                If WProceso = 1 Then
                    Call Revalida_Click
                        Else
                    Call Bloqueo_Click
                End If
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim VectorCosto(120, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    VectorCosto(1, 1) = Producto
    VectorCosto(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If VectorCosto(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + VectorCosto(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    VectorCosto(Lugar, 1) = Articulo2
                                    VectorCosto(Lugar, 2) = Str$(Cantidad * Val(VectorCosto(Cicla, 2)))
                                End If
                            Case "M"
                     Rem BY NAN
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = VectorCosto(Cicla, 2)
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
            
            If Entra = "S" Then
                If Left$(VectorCosto(Cicla, 1), 2) <> "PT" Then
                    Renglon = Renglon + 1
                    Auxiliar(Renglon, 1) = Left$(VectorCosto(Cicla, 1), 3) + Right$(VectorCosto(Cicla, 1), 7)
                    Auxiliar(Renglon, 2) = 1
                    Auxiliar(Renglon, 3) = VectorCosto(Cicla, 2)
                End If
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For DA = 1 To Renglon
        Articulo = Auxiliar(DA, 1)
        Cantidad = Auxiliar(DA, 2)
        WVectorCosto = Auxiliar(DA, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(WVectorCosto))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(WVectorCosto))
            rstArticulo.Close
        End If
    Next DA

End Sub




Public Sub SendEmail()

    Dim objOutlook As Object
    Dim objMailItem

    Dim NumOfPath As Integer, i As Integer
    Dim AtachPath As String

    On Error GoTo 10

    NumOfPath = 0
    AllPath = ""
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
        .to = EmailAddress
        .cc = CopiaAddress
        .Subject = MSubject
        .Body = MBody
        Rem .Attachments.Add MAttach
        Rem If MAttachI <> "" Then
        Rem     .Attachments.Add MAttachI
        Rem End If
        Rem If MAttachII <> "" Then
        Rem     .Attachments.Add MAttachII
        Rem End If
        Rem If MAttachIII > "" Then
        Rem     .Attachments.Add MAttachIII
        Rem End If
        Rem If MAttachIV <> "" Then
        Rem     .Attachments.Add MAttachIV
        Rem End If
        Rem If MAttachV <> "" Then
        Rem     .Attachments.Add MAttachV
        Rem End If
        Rem If MAttachVI <> "" Then
        Rem     .Attachments.Add MAttachVI
        Rem End If
        Rem If MAttachVII <> "" Then
        Rem     .Attachments.Add MAttachVII
        Rem End If
        Rem If MAttachVIII <> "" Then
        Rem     .Attachments.Add MAttachVIII
        Rem End If
        .Send
    End With

    Set objMailItem = Nothing
    Set objOutlook = Nothing
            
    Exit Sub

exit10:
    Exit Sub

10:
    If Err.Number = 429 Then
        MsgBox "Error on connecting with Outlook"
            Else
        MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    End If
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    AllPath = ""

    Resume exit10

End Sub




Private Sub Limpia_Vector()

    WGrilla.Clear
    WGrilla.Font.Bold = True
    
    WGrilla.FixedCols = 1
    WGrilla.Cols = 6
    WGrilla.FixedRows = 1
    WGrilla.Rows = 51
    
    WGrilla.ColWidth(0) = 200
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        Select Case Ciclo
            Case 1
                WGrilla.Text = "Tipo"
                WGrilla.ColWidth(Ciclo) = 500
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WGrilla.Text = "Materia Prima"
                WGrilla.ColWidth(Ciclo) = 1400
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WGrilla.Text = "Producto Terminado"
                WGrilla.ColWidth(Ciclo) = 1400
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                WGrilla.Text = "Descripcion"
                WGrilla.ColWidth(Ciclo) = 4200
                WGrilla.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 5
                WGrilla.Text = "Cantidad"
                WGrilla.ColWidth(Ciclo) = 1200
                WGrilla.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WGrilla.Row = 0
    For Ciclo = 1 To WGrilla.Cols - 1
        WGrilla.Col = Ciclo
        WTitulo(Ciclo).Text = WGrilla.Text
        WTitulo(Ciclo).Left = WGrilla.CellLeft + WGrilla.Left
        WTitulo(Ciclo).Top = WGrilla.CellTop + WGrilla.Top
        WTitulo(Ciclo).Width = WGrilla.CellWidth
        WTitulo(Ciclo).Height = WGrilla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WGrilla.Cols - 1
        WAncho = WAncho + WGrilla.ColWidth(Ciclo)
    Next Ciclo
    WGrilla.Width = WAncho

    ' Size the columns.
    Font.Name = WGrilla.Font.Name
    Font.Size = WGrilla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WGrilla.AllowUserResizing = flexResizeBoth
    
    WGrilla.Visible = True
    
    WGrilla.Col = 1
    WGrilla.Row = 1
    
End Sub







Private Sub Limpia_VectorII()

    WGrillaExplo.Clear
    WGrillaExplo.Font.Bold = True
    
    WGrillaExplo.FixedCols = 1
    WGrillaExplo.Cols = 4
    WGrillaExplo.FixedRows = 1
    WGrillaExplo.Rows = 51
    
    WGrillaExplo.ColWidth(0) = 200
    WGrillaExplo.Row = 0
    For Ciclo = 1 To WGrillaExplo.Cols - 1
        WGrillaExplo.Col = Ciclo
        Select Case Ciclo
            Case 1
                WGrillaExplo.Text = "Materia Prima"
                WGrillaExplo.ColWidth(Ciclo) = 1400
                WGrillaExplo.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WGrillaExplo.Text = "Descripcion"
                WGrillaExplo.ColWidth(Ciclo) = 4200
                WGrillaExplo.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WGrillaExplo.Text = "Cantidad"
                WGrillaExplo.ColWidth(Ciclo) = 1200
                WGrillaExplo.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WGrillaExplo.Cols - 1
        WAncho = WAncho + WGrillaExplo.ColWidth(Ciclo)
    Next Ciclo
    WGrillaExplo.Width = WAncho

    ' Size the columns.
    Font.Name = WGrillaExplo.Font.Name
    Font.Size = WGrillaExplo.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WGrillaExplo.AllowUserResizing = flexResizeBoth
    
    WGrillaExplo.Visible = True
    
    WGrillaExplo.Col = 1
    WGrillaExplo.Row = 1
    
End Sub



