VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaIIIProduccion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Instrucciones de Produccion de P.T."
   ClientHeight    =   8520
   ClientLeft      =   180
   ClientTop       =   285
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.CommandButton GrabaRegistro 
      Caption         =   "Trae Registro Version"
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
      Left            =   7320
      TabIndex        =   62
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Frame PantaGraba 
      Height          =   2895
      Left            =   3000
      TabIndex        =   54
      Top             =   2160
      Visible         =   0   'False
      Width           =   5535
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
         HideSelection   =   0   'False
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   60
         Text            =   " "
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox PasoII 
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
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   55
         Text            =   " "
         Top             =   840
         Width           =   495
      End
      Begin MSMask.MaskEdBox TerminadoII 
         Height          =   285
         Left            =   1080
         TabIndex        =   56
         Top             =   360
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
      Begin VB.Label Label12 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image GrabaII 
         Height          =   480
         Left            =   1440
         MouseIcon       =   "cargaiiiproduccion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "cargaiiiproduccion.frx":030A
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label14 
         Caption         =   "Etapa"
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
         TabIndex        =   59
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label DesTerminadoII 
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
         Left            =   2640
         TabIndex        =   57
         Top             =   360
         Width           =   2535
      End
      Begin VB.Image CierraII 
         Height          =   480
         Left            =   3000
         MouseIcon       =   "cargaiiiproduccion.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "cargaiiiproduccion.frx":0E56
         ToolTipText     =   "Salida"
         Top             =   1920
         Width           =   480
      End
   End
   Begin VB.CommandButton GrabaVersion 
      Caption         =   "Graba Version"
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
      Left            =   10560
      TabIndex        =   53
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox ControlCambio 
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
      MaxLength       =   100
      TabIndex        =   51
      Top             =   5760
      Width           =   8415
   End
   Begin VB.TextBox Version 
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
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   49
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11040
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Metodo 
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
      MaxLength       =   4
      TabIndex        =   45
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Limpieza 
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
      Left            =   6840
      TabIndex        =   43
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox Libera 
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
      Left            =   1800
      TabIndex        =   41
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton AltaEpata 
      Caption         =   "Inserta Etapa"
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
      Left            =   9720
      TabIndex        =   40
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton AltaPoe 
      Height          =   735
      Left            =   9000
      MouseIcon       =   "cargaiiiproduccion.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Ingresa POE"
      Top             =   6840
      Width           =   615
   End
   Begin VB.Frame IngresaPoe 
      BackColor       =   &H00C0FFFF&
      Height          =   3495
      Left            =   2160
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   3975
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Image MuestraFoto 
         Height          =   2295
         Left            =   6480
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.ComboBox Humedad 
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
      Left            =   9240
      TabIndex        =   33
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Epp 
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
      Left            =   4920
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Peso 
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
      Left            =   9840
      TabIndex        =   29
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Equipo 
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
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   480
      Width           =   855
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   4455
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   7858
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso de Procedimiento"
      TabPicture(0)   =   "cargaiiiproduccion.frx":226C
      Tab(0).ControlCount=   10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WVector1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "WTitulo(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "WTitulo(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "WTitulo(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "WTexto1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "WCombo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "WTexto2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "WTitulo(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "WTitulo(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "WTexto3"
      Tab(0).Control(9).Enabled=   0   'False
      TabCaption(1)   =   "Ingreso de Controles de Calidad"
      TabPicture(1)   =   "cargaiiiproduccion.frx":2288
      Tab(1).ControlCount=   5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WCombo12"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "WTexto32"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto22"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "WTexto12"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "WVector2"
      Tab(1).Control(4).Enabled=   0   'False
      Begin VB.ComboBox WCombo12 
         Height          =   315
         Left            =   -71160
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -72000
         TabIndex        =   23
         Top             =   1440
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
      Begin VB.TextBox WTexto22 
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
         Left            =   -72720
         TabIndex        =   22
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox WTexto12 
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
         Left            =   -73440
         TabIndex        =   21
         Top             =   1440
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Top             =   1380
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1920
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1920
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
         Left            =   3000
         TabIndex        =   15
         Top             =   1380
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         Top             =   1320
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
         Left            =   2400
         TabIndex        =   13
         Top             =   1380
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   1920
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
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   1860
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
         Index           =   5
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   " "
         Top             =   1920
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3975
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7011
         _Version        =   327680
         BackColor       =   16777152
      End
   End
   Begin VB.TextBox Paso 
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
      TabIndex        =   8
      Text            =   " "
      Top             =   120
      Width           =   495
   End
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
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   1740
      Left            =   2280
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   6480
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
      ItemData        =   "cargaiiiproduccion.frx":22A4
      Left            =   120
      List            =   "cargaiiiproduccion.frx":22AB
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
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
   Begin MSMask.MaskEdBox FechaVersion 
      Height          =   285
      Left            =   7800
      TabIndex        =   50
      Top             =   120
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
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6480
      TabIndex        =   48
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Metodo"
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
      TabIndex        =   46
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Limpieza Equipo"
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
      Left            =   5040
      TabIndex        =   44
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Libera Area"
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
      TabIndex        =   42
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8280
      MouseIcon       =   "cargaiiiproduccion.frx":22B9
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":25C3
      ToolTipText     =   "Elimina el Registro"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Contr. Humedad"
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
      Left            =   7680
      TabIndex        =   34
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Epp"
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
      Left            =   4440
      TabIndex        =   32
      Top             =   480
      Width           =   735
   End
   Begin VB.Label DesEpp 
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
      Left            =   5880
      TabIndex        =   31
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image AltaRenglon 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "cargaiiiproduccion.frx":2E05
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":310F
      ToolTipText     =   "Agrega Renglon"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Peso"
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
      Left            =   9240
      TabIndex        =   28
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label DesEquipo 
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
      Left            =   2040
      TabIndex        =   27
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   26
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Etapa"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesTerminado 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   9960
      MouseIcon       =   "cargaiiiproduccion.frx":3551
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":385B
      ToolTipText     =   "Salida"
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "cargaiiiproduccion.frx":409D
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":43A7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   9120
      MouseIcon       =   "cargaiiiproduccion.frx":4BE9
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":4EF3
      ToolTipText     =   "Consulta de Datos"
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8280
      MouseIcon       =   "cargaiiiproduccion.frx":5735
      MousePointer    =   99  'Custom
      Picture         =   "cargaiiiproduccion.frx":5A3F
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6240
      Width           =   480
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "PrgCargaIIIProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayos As Recordset
Dim spEnsayos As String
Dim rstCargaIII As Recordset
Dim spCargaIII As String
Dim rstCargaV As Recordset
Dim spCargaV As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer

Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Dim XPaso As String
Dim Renglon As Integer
Dim ZEnsayo As String
Dim ZValor As String

Dim WPasaVector(5000, 26) As String
Dim ZZPasa(1000, 30) As String

Dim ZClave As String
Dim ZTerminado As String
Dim ZPaso As String
Dim ZRenglon As String
Dim ZArticulo As String
Dim ZPTerminado As String
Dim ZLetra As String
Dim ZDescripcion As String
Dim ZCantidad As String
Dim ZCantidadII As String
Dim ZPartida As String
Dim ZCantidadPartida As String
Dim ZEquipo As String
Dim ZPeso As String
Dim ZTipo As String
Dim ZItem As String
Dim ZEpp As String
Dim ZDesEpp As String
Dim ZCorteItem As String
Dim ZImprePeso As String
Dim ZHumedad As String
Dim ZImpreHumedad As String
Dim ZLibera As String
Dim ZLimpieza As String
Dim ZMetodo As String

Dim ZZZArticulo As String
Dim ZZZPTerminado As String
Dim ZZZLetra As String
Dim ZZZDescripcion As String
Dim ZZZCantidad As Double

Dim ZZZEnsayo As Integer
Dim ZZZValor As String


Dim ZZZClave As String
Dim ZZZTerminado As String
Dim ZZZPaso As Integer
Dim ZZZRenglon As Integer

Dim ZZZEquipo As Integer
Dim ZZZPeso As Double
Dim ZZZTipo As String
Dim ZZZItem As String
Dim ZZZEpp As Double
Dim ZZZDesEpp As String
Dim ZZZCorteItem As Integer
Dim ZZZImprePeso As String
Dim ZZZHumedad As Double
Dim ZZZImpreHumedad As String
Dim ZZZLibera As String
Dim ZZZLimpieza As String
Dim ZZZMetodo As String


Dim ZZTerminado As String
Dim ZZPaso As String
Dim ZZRenglon As String
Dim ZZArticulo As String
Dim ZZPTerminado As String
Dim ZZLetra As String
Dim ZZDescripcion As String
Dim ZZCantidad As String
Dim ZZCantidadII As String
Dim ZZPartida As String
Dim ZZCantidadPartida As String
Dim ZZEquipo As String
Dim ZZPeso As String
Dim ZZTipo As String
Dim ZZItem As String
Dim ZZEpp As String
Dim ZZDesEpp As String
Dim ZZCorteItem As String
Dim ZZImprePeso As String
Dim ZZHumedad As String
Dim ZZImpreHumedad As String
Dim ZZLibera As String
Dim ZZLimpieza As String
Dim ZZMetodo As String
Dim ZZControlCambio As String
Dim ZZVersion As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Rem para el vector II

Dim WBorraII(1000, 20) As String
Dim WParametrosII(10, 20) As Double
Dim WFormatoII(20) As String
Dim WControlII As String

Private Sub AltaEpata_Click()

    T$ = "Registro de Produccion"
    m$ = "Desea Insertar una etapa"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        ZLugar = 0
        Erase WPasaVector
        
        ZSql = " "
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaIII"
        ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and CargaIII.Paso >= " + "'" + Paso.Text + "'"
        ZSql = ZSql + " and CargaIII.Paso <> 99"
        ZSql = ZSql + " Order by CargaIII.Clave"
    
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIII.RecordCount > 0 Then
            With rstCargaIII
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        
                        WPasaVector(ZLugar, 1) = rstCargaIII!Clave
                        WPasaVector(ZLugar, 2) = rstCargaIII!Terminado
                        WPasaVector(ZLugar, 3) = Str$(rstCargaIII!Paso)
                        WPasaVector(ZLugar, 4) = Str$(rstCargaIII!Renglon)
                        WPasaVector(ZLugar, 5) = rstCargaIII!Articulo
                        WPasaVector(ZLugar, 6) = rstCargaIII!PTerminado
                        WPasaVector(ZLugar, 7) = rstCargaIII!Letra
                        WPasaVector(ZLugar, 8) = rstCargaIII!Descripcion
                        WPasaVector(ZLugar, 9) = Str$(rstCargaIII!Cantidad)
                        WPasaVector(ZLugar, 10) = ""
                        WPasaVector(ZLugar, 11) = ""
                        WPasaVector(ZLugar, 12) = ""
                        ZZEquipo = IIf(IsNull(rstCargaIII!Equipo), "", rstCargaIII!Equipo)
                        WPasaVector(ZLugar, 13) = ZZEquipo
                        ZZPeso = IIf(IsNull(rstCargaIII!Peso), "", rstCargaIII!Peso)
                        WPasaVector(ZLugar, 14) = ZZPeso
                        ZZTipo = IIf(IsNull(rstCargaIII!Tipo), "", rstCargaIII!Tipo)
                        WPasaVector(ZLugar, 15) = ZZTipo
                        ZZItem = IIf(IsNull(rstCargaIII!Item), "", rstCargaIII!Item)
                        WPasaVector(ZLugar, 16) = ZZItem
                        ZZEpp = IIf(IsNull(rstCargaIII!Epp), "", rstCargaIII!Epp)
                        WPasaVector(ZLugar, 17) = ZZEpp
                        ZZDesEpp = IIf(IsNull(rstCargaIII!DesEpp), "", rstCargaIII!DesEpp)
                        WPasaVector(ZLugar, 18) = ZZDesEpp
                        ZZCorteItem = IIf(IsNull(rstCargaIII!CorteItem), "", rstCargaIII!CorteItem)
                        WPasaVector(ZLugar, 19) = ZZCorteItem
                        ZZImprePeso = IIf(IsNull(rstCargaIII!ImprePeso), "", rstCargaIII!ImprePeso)
                        WPasaVector(ZLugar, 20) = ZZImprePeso
                        ZZHumedad = IIf(IsNull(rstCargaIII!Humedad), "", rstCargaIII!Humedad)
                        WPasaVector(ZLugar, 21) = ZZHumedad
                        ZZImpreHumedad = IIf(IsNull(rstCargaIII!ImpreHumedad), "", rstCargaIII!ImpreHumedad)
                        WPasaVector(ZLugar, 22) = ZZImpreHumedad
                        ZZLibera = IIf(IsNull(rstCargaIII!Libera), "", rstCargaIII!Libera)
                        WPasaVector(ZLugar, 23) = ZZLibera
                        ZZLimpieza = IIf(IsNull(rstCargaIII!Limpieza), "", rstCargaIII!Limpieza)
                        WPasaVector(ZLugar, 24) = ZZLimpieza
                        ZZMetodo = IIf(IsNull(rstCargaIII!Metodo), "", rstCargaIII!Metodo)
                        WPasaVector(ZLugar, 25) = ZZMetodo
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaIII.Close
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaIII"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and Paso >= " + "'" + Paso.Text + "'"
        ZSql = ZSql + " and Paso <> 99"
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        ZPasa = 0
        ZCorteEtapaI = ""
        ZCorteEtapaII = ""
        
        For Ciclo = 1 To ZLugar
        
            ZTerminado = WPasaVector(Ciclo, 2)
            
            ZPaso = Str$(Val(WPasaVector(Ciclo, 3)) + 1)
            Auxi = ZPaso
            Call Ceros(Auxi, 4)
            
            ZRenglon = WPasaVector(Ciclo, 4)
            Auxi1 = ZRenglon
            Call Ceros(Auxi1, 2)
            
            ZClave = ZTerminado + Auxi + Auxi1
            
            ZArticulo = WPasaVector(Ciclo, 5)
            ZPTerminado = WPasaVector(Ciclo, 6)
            ZLetra = WPasaVector(Ciclo, 7)
            ZDescripcion = WPasaVector(Ciclo, 8)
            ZCantidad = WPasaVector(Ciclo, 9)
            ZCantidadII = WPasaVector(Ciclo, 10)
            ZPartida = WPasaVector(Ciclo, 11)
            ZCantidadPartida = WPasaVector(Ciclo, 12)
            ZEquipo = WPasaVector(Ciclo, 13)
            ZPeso = WPasaVector(Ciclo, 14)
            ZTipo = WPasaVector(Ciclo, 15)
            Rem ZItem = WPasaVector(Ciclo, 16)
            ZEpp = WPasaVector(Ciclo, 17)
            ZDesEpp = WPasaVector(Ciclo, 18)
            ZCorteItem = WPasaVector(Ciclo, 19)
            ZImprePeso = WPasaVector(Ciclo, 20)
            ZHumedad = WPasaVector(Ciclo, 21)
            ZImpreHumedad = WPasaVector(Ciclo, 22)
            ZLibera = WPasaVector(Ciclo, 23)
            ZLimpieza = WPasaVector(Ciclo, 24)
            ZMetodo = WPasaVector(Ciclo, 25)
            ZItem = ""
            
            If ZPasa = 0 Then
                ZPasa = 1
                ZCorteEtapaI = ZPaso
                ZCorteEtapaII = ZCorteItem
                ZItem = Trim(Str$(Val(ZPaso))) + "." + Trim(Str$(ZCorteItem))
            End If
            
            If ZCorteEtapaI <> ZPaso Or ZCorteEtapaII <> ZCorteItem Then
                ZCorteEtapaI = ZPaso
                ZCorteEtapaII = ZCorteItem
                ZItem = Trim(Str$(Val(ZPaso))) + "." + Trim(Str$(ZCorteItem))
            End If

            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaIII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Paso ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "PTerminado ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "Peso ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Epp ,"
            ZSql = ZSql + "DesEpp ,"
            ZSql = ZSql + "CorteItem ,"
            ZSql = ZSql + "ImprePeso ,"
            ZSql = ZSql + "Libera ,"
            ZSql = ZSql + "Limpieza ,"
            ZSql = ZSql + "Metodo ,"
            ZSql = ZSql + "Humedad ,"
            ZSql = ZSql + "ImpreHumedad )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + ZPaso + "',"
            ZSql = ZSql + "'" + ZRenglon + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZPTerminado + "',"
            ZSql = ZSql + "'" + ZLetra + "',"
            ZSql = ZSql + "'" + ZDescripcion + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZEquipo + "',"
            ZSql = ZSql + "'" + ZPeso + "',"
            ZSql = ZSql + "'" + ZTipo + "',"
            ZSql = ZSql + "'" + ZItem + "',"
            ZSql = ZSql + "'" + ZEpp + "',"
            ZSql = ZSql + "'" + ZDesEpp + "',"
            ZSql = ZSql + "'" + ZCorteItem + "',"
            ZSql = ZSql + "'" + ZImprePeso + "',"
            ZSql = ZSql + "'" + ZLibera + "',"
            ZSql = ZSql + "'" + ZLimpieza + "',"
            ZSql = ZSql + "'" + ZMetodo + "',"
            ZSql = ZSql + "'" + ZHumedad + "',"
            ZSql = ZSql + "'" + ZImpreHumedad + "')"
                
            rsCargaIII = ZSql
            Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Ciclo
        
    End If
    
    Call Limpia_Click

End Sub

Private Sub AltaRenglon_Click()

    RenglonAuxiliar = WVector1.Row

    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To 199
        WVector1.Row = Ciclo
        EntraVector = EntraVector + 1
        For Ciclo1 = 1 To WVector1.Cols - 1
            WVector1.Col = Ciclo1
            WBorra(EntraVector, Ciclo1) = WVector1.Text
        Next Ciclo1
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        If Ciclo >= RenglonAuxiliar Then
            ZLugar = Ciclo + 1
                Else
            ZLugar = Ciclo
        End If
        WVector1.Row = ZLugar
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub Command1_Click()
    
            ZZPalabraI = "F. Vencimineto"
            ZZPalabraII = "F. Reanalisis "
    
            Sql1 = "Select *"
            Sql2 = " FROM CargaIII"
            Sql3 = " Order by Clave"
            spCargaIII = Sql1 + Sql2 + Sql3
            Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaIII.RecordCount > 0 Then
                With rstCargaIII
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            da = Len(rstTerminado!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
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
                rstCargaIII.Close
            End If
    
    
    
    
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE CargaIII SET "
    Rem ZSql = ZSql + " Version = " + "'" + "1" + "',"
    Rem ZSql = ZSql + " FechaVersion = " + "'" + "01/01/2010" + "',"
    Rem ZSql = ZSql + " OrdFechaVersion = " + "'" + "20100101" + "'"
    Rem spCargaIII = ZSql
    Rem Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
End Sub

Private Sub cmdDelete_Click()

    T$ = "Registro de Produccion"
    m$ = "Desea eliminar la etapa"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Sql1 = "DELETE CargaIII"
        Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
        Sql3 = " and Paso = " + "'" + Paso.Text + "'"
        rsCargaIII = Sql1 + Sql2 + Sql3
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        ZLugar = 0
        Erase WPasaVector
        
        ZSql = " "
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaIII"
        ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and CargaIII.Paso > " + "'" + Paso.Text + "'"
        ZSql = ZSql + " and CargaIII.Paso <> 99"
        ZSql = ZSql + " Order by CargaIII.Clave"
    
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIII.RecordCount > 0 Then
            With rstCargaIII
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        
                        ZZZClave = IIf(IsNull(rstCargaIII!Clave), "", rstCargaIII!Clave)
                        WPasaVector(ZLugar, 1) = ZZZClave
                        
                        ZZZTerminado = IIf(IsNull(rstCargaIII!Terminado), "", rstCargaIII!Terminado)
                        WPasaVector(ZLugar, 2) = ZZZTerminado
                        
                        ZZZPaso = IIf(IsNull(rstCargaIII!Paso), "0", rstCargaIII!Paso)
                        WPasaVector(ZLugar, 3) = Str$(ZZZPaso)
                        
                        ZZZRenglon = IIf(IsNull(rstCargaIII!Renglon), "0", rstCargaIII!Renglon)
                        WPasaVector(ZLugar, 4) = Str$(ZZZRenglon)
                        
                        ZZZArticulo = IIf(IsNull(rstCargaIII!Articulo), "", rstCargaIII!Articulo)
                        WPasaVector(ZLugar, 5) = ZZZArticulo
                        
                        ZZZPTerminado = IIf(IsNull(rstCargaIII!PTerminado), "", rstCargaIII!PTerminado)
                        WPasaVector(ZLugar, 6) = ZZZPTerminado
                        
                        ZZZLetra = IIf(IsNull(rstCargaIII!Letra), "", rstCargaIII!Letra)
                        WPasaVector(ZLugar, 7) = ZZZLetra
                        
                        ZZZDescripcion = IIf(IsNull(rstCargaIII!Descripcion), "", rstCargaIII!Descripcion)
                        WPasaVector(ZLugar, 8) = ZZZDescripcion
                        
                        ZZZCantidad = IIf(IsNull(rstCargaIII!Cantidad), "0", rstCargaIII!Cantidad)
                        WPasaVector(ZLugar, 9) = Str$(ZZZCantidad)
                        
                        WPasaVector(ZLugar, 10) = ""
                        WPasaVector(ZLugar, 11) = ""
                        WPasaVector(ZLugar, 12) = ""
                        
                        ZZZEquipo = IIf(IsNull(rstCargaIII!Equipo), "0", rstCargaIII!Equipo)
                        WPasaVector(ZLugar, 13) = Str$(ZZZEquipo)
                        
                        ZZZPeso = IIf(IsNull(rstCargaIII!Peso), "0", rstCargaIII!Peso)
                        WPasaVector(ZLugar, 14) = Str$(ZZZPeso)
                        
                        ZZZTipo = IIf(IsNull(rstCargaIII!Tipo), "", rstCargaIII!Tipo)
                        WPasaVector(ZLugar, 15) = ZZZTipo
                        
                        ZZZItem = IIf(IsNull(rstCargaIII!Item), "", rstCargaIII!Item)
                        WPasaVector(ZLugar, 16) = ZZZItem
                        
                        ZZZEpp = IIf(IsNull(rstCargaIII!Epp), "0", rstCargaIII!Epp)
                        WPasaVector(ZLugar, 17) = Str$(ZZZEpp)
                        
                        ZZZDesEpp = IIf(IsNull(rstCargaIII!DesEpp), "", rstCargaIII!DesEpp)
                        WPasaVector(ZLugar, 18) = ZZZDesEpp
                        
                        ZZZCorteItem = IIf(IsNull(rstCargaIII!CorteItem), "0", rstCargaIII!CorteItem)
                        WPasaVector(ZLugar, 19) = Str$(ZZZCorteItem)
                        
                        ZZZImprePeso = IIf(IsNull(rstCargaIII!ImprePeso), "", rstCargaIII!ImprePeso)
                        WPasaVector(ZLugar, 20) = ZZZImprePeso
                        
                        ZZZHumedad = IIf(IsNull(rstCargaIII!Humedad), "0", rstCargaIII!Humedad)
                        WPasaVector(ZLugar, 21) = Str$(ZZZHumedad)
                        
                        ZZZImpreHumedad = IIf(IsNull(rstCargaIII!ImpreHumedad), "", rstCargaIII!ImpreHumedad)
                        WPasaVector(ZLugar, 22) = ZZZImpreHumedad
                        
                        ZZZLibera = IIf(IsNull(rstCargaIII!Libera), "", rstCargaIII!Libera)
                        WPasaVector(ZLugar, 23) = ZZZLibera
                        
                        ZZZLimpieza = IIf(IsNull(rstCargaIII!Limpieza), "", rstCargaIII!Limpieza)
                        WPasaVector(ZLugar, 24) = ZZZLimpieza
                        
                        ZZZMetodo = IIf(IsNull(rstCargaIII!Metodo), "", rstCargaIII!Metodo)
                        WPasaVector(ZLugar, 25) = ZZZMetodo
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCargaIII.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "DELETE CargaIII"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and Paso > " + "'" + Paso.Text + "'"
        ZSql = ZSql + " and Paso <> 99"
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        ZPasa = 0
        ZCorteEtapaI = ""
        ZCorteEtapaII = ""
        
        For Ciclo = 1 To ZLugar
        
            ZTerminado = WPasaVector(Ciclo, 2)
            
            ZPaso = Str$(Val(WPasaVector(Ciclo, 3)) - 1)
            Auxi = ZPaso
            Call Ceros(Auxi, 4)
            
            ZRenglon = WPasaVector(Ciclo, 4)
            Auxi1 = ZRenglon
            Call Ceros(Auxi1, 2)
            
            ZClave = ZTerminado + Auxi + Auxi1
            
            ZArticulo = WPasaVector(Ciclo, 5)
            ZPTerminado = WPasaVector(Ciclo, 6)
            ZLetra = WPasaVector(Ciclo, 7)
            ZDescripcion = WPasaVector(Ciclo, 8)
            ZCantidad = WPasaVector(Ciclo, 9)
            ZCantidadII = WPasaVector(Ciclo, 10)
            ZPartida = WPasaVector(Ciclo, 11)
            ZCantidadPartida = WPasaVector(Ciclo, 12)
            ZEquipo = WPasaVector(Ciclo, 13)
            ZPeso = WPasaVector(Ciclo, 14)
            ZTipo = WPasaVector(Ciclo, 15)
            Rem ZItem = WPasaVector(Ciclo, 16)
            ZEpp = WPasaVector(Ciclo, 17)
            ZDesEpp = WPasaVector(Ciclo, 18)
            ZCorteItem = WPasaVector(Ciclo, 19)
            ZImprePeso = WPasaVector(Ciclo, 20)
            ZHumedad = WPasaVector(Ciclo, 21)
            ZImpreHumedad = WPasaVector(Ciclo, 22)
            ZLibera = WPasaVector(Ciclo, 23)
            ZLimpieza = WPasaVector(Ciclo, 24)
            ZMetodo = WPasaVector(Ciclo, 25)
            ZItem = ""
            
            If ZPasa = 0 Then
                ZPasa = 1
                ZCorteEtapaI = ZPaso
                ZCorteEtapaII = ZCorteItem
                ZItem = Trim(Str$(Val(ZPaso))) + "." + Trim(Str$(ZCorteItem))
            End If
            
            If ZCorteEtapaI <> ZPaso Or ZCorteEtapaII <> ZCorteItem Then
                ZCorteEtapaI = ZPaso
                ZCorteEtapaII = ZCorteItem
                ZItem = Trim(Str$(Val(ZPaso))) + "." + Trim(Str$(ZCorteItem))
            End If
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaIII ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Paso ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "PTerminado ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "Peso ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Epp ,"
            ZSql = ZSql + "DesEpp ,"
            ZSql = ZSql + "CorteItem ,"
            ZSql = ZSql + "ImprePeso ,"
            ZSql = ZSql + "Libera ,"
            ZSql = ZSql + "Limpieza ,"
            ZSql = ZSql + "Metodo ,"
            ZSql = ZSql + "Humedad ,"
            ZSql = ZSql + "ImpreHumedad )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + ZTerminado + "',"
            ZSql = ZSql + "'" + ZPaso + "',"
            ZSql = ZSql + "'" + ZRenglon + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZPTerminado + "',"
            ZSql = ZSql + "'" + ZLetra + "',"
            ZSql = ZSql + "'" + ZDescripcion + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZEquipo + "',"
            ZSql = ZSql + "'" + ZPeso + "',"
            ZSql = ZSql + "'" + ZTipo + "',"
            ZSql = ZSql + "'" + ZItem + "',"
            ZSql = ZSql + "'" + ZEpp + "',"
            ZSql = ZSql + "'" + ZDesEpp + "',"
            ZSql = ZSql + "'" + ZCorteItem + "',"
            ZSql = ZSql + "'" + ZImprePeso + "',"
            ZSql = ZSql + "'" + ZLibera + "',"
            ZSql = ZSql + "'" + ZLimpieza + "',"
            ZSql = ZSql + "'" + ZMetodo + "',"
            ZSql = ZSql + "'" + ZHumedad + "',"
            ZSql = ZSql + "'" + ZImpreHumedad + "')"
                
            rsCargaIII = ZSql
            Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
            
        Next Ciclo
        
    End If
    
    Call Limpia_Click

End Sub

Private Sub Command2_Click()

    Dim ZZVector(1000) As String
    
    Erase ZZVector
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Renglon = " + "'" + "1" + "'"
    ZSql = ZSql + " and CargaIII.Paso = " + "'" + "1" + "'"
    ZSql = ZSql + " Order by CargaIII.Version"
    
    rsCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        With rstCargaIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZVersion = IIf(IsNull(rstCargaIII!Version), "0", rstCargaIII!Version)
                    
                    If ZZVersion = 0 Then
                        ZZLugar = ZZLugar + 1
                        ZZVector(ZZLugar) = rstCargaIII!Terminado
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIII.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZTerminado = ZZVector(Ciclo)
        ZZVersion = ""
        ZZFechaVersion = "  /  /    "
        ZZOrdFechaVersion = ""
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaIIIVersion"
        ZSql = ZSql + " Where CargaIIIVersion.Terminado = " + "'" + ZZTerminado + "'"
        ZSql = ZSql + " Order by CargaIIIVersion.Version Desc"
        
        rsCargaIIIVersion = ZSql
        Set rstCargaIIIVersion = db.OpenRecordset(rsCargaIIIVersion, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIIIVersion.RecordCount > 0 Then
            With rstCargaIIIVersion
                .MoveFirst
                ZZVersion = IIf(IsNull(rstCargaIIIVersion!Version), "0", rstCargaIIIVersion!Version)
                ZZFechaVersion = IIf(IsNull(rstCargaIIIVersion!FechaVersionII), "  /  /    ", rstCargaIIIVersion!FechaVersionII)
                ZZOrdFechaVersion = IIf(IsNull(rstCargaIIIVersion!OrdFechaVersionII), "", rstCargaIIIVersion!OrdFechaVersionII)
            End With
            rstCargaIIIVersion.Close
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIII SET "
        ZSql = ZSql + " Version = " + "'" + ZZVersion + "',"
        ZSql = ZSql + " FechaVersion = " + "'" + ZZFechaVersion + "',"
        ZSql = ZSql + " OrdFechaVersion = " + "'" + ZZOrdFechaVersion + "'"
        ZSql = ZSql + " Where Terminado = " + "'" + ZZTerminado + "'"
        spCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
    
    Next Ciclo

Stop


End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"
     Opcion.AddItem "Ensayos"
     Opcion.AddItem "Equipos"
     Opcion.AddItem "Epp"
     Opcion.AddItem "Texto Fijo"

     Opcion.Visible = True
     
End Sub


Private Sub GrabaVersion_Click()
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " Version = " + "'" + Version.Text + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)

    m$ = "Operacion Exitosa"
    A% = MsgBox(m$, 0, "Carga de Registro de Produccion")

End Sub

Private Sub Peso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Equipo.SetFocus
    End If
End Sub

Private Sub Equipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Equipo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Equipo"
            Sql3 = " Where Equipo.Codigo = " + "'" + Equipo.Text + "'"
            spEquipo = Sql1 + Sql2 + Sql3
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                DesEquipo.Caption = rstEquipo!Descripcion
                rstEquipo.Close
                Epp.SetFocus
            End If
                Else
            DesEquipo.Caption = ""
            Epp.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Equipo.Text = ""
        DesEquipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Epp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Epp.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM MaterialAuxiliar"
            Sql3 = " Where MaterialAuxiliar.Codigo = " + "'" + Epp.Text + "'"
            spMaterialAuxiliar = Sql1 + Sql2 + Sql3
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                DesEpp.Caption = rstMaterialAuxiliar!Descripcion
                rstMaterialAuxiliar.Close
                Humedad.SetFocus
            End If
                Else
            Humedad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Epp.Text = ""
        DesEpp.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Humedad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Libera.SetFocus
    End If
End Sub

Private Sub Libera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Limpieza.SetFocus
    End If
End Sub

Private Sub Limpieza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Metodo.SetFocus
    End If
End Sub

Private Sub Metodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tablas.Tab = 0
        WVector1.TopRow = 1
        WVector1.Col = 1
        WVector1.Row = 1
        WVector2.TopRow = 1
        WVector2.Col = 1
        WVector2.Row = 1
        Call StartEdit
    End If
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
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
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
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
            End If
            
        Case 3
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 9
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
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayos!Codigo) + " " + rstEnsayos!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayos!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
    
            Call Conecta_Empresa
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Equipo"
            ZSql = ZSql + " Order by Codigo"
            spEquipo = ZSql
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                With rstEquipo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEquipo!Codigo) + " " + rstEquipo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstEquipo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEquipo.Close
            End If
            
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM MaterialAuxiliar"
            ZSql = ZSql + " Order by Codigo"
            spMaterialAuxiliar = ZSql
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                With rstMaterialAuxiliar
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstMaterialAuxiliar!Codigo) + " " + rstMaterialAuxiliar!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstMaterialAuxiliar!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMaterialAuxiliar.Close
            End If
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TextoFijo"
            ZSql = ZSql + " Order by Codigo"
            spTextoFijo = ZSql
            Set rstTextoFijo = db.OpenRecordset(spTextoFijo, dbOpenSnapshot, dbSQLPassThrough)
            If rstTextoFijo.RecordCount > 0 Then
                With rstTextoFijo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstTextoFijo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTextoFijo!Descripcion
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTextoFijo.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgCargaIIIProduccion.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    If Humedad.ListIndex = 2 Then
        m$ = "Se debe definir si se debe tomar valores de humedad"
        A% = MsgBox(m$, 0, "Carga de Registro de Produccion")
        Exit Sub
    End If

    If Peso.ListIndex = 2 Then
        m$ = "Se debe definir si se debe tomar valores de peso"
        A% = MsgBox(m$, 0, "Carga de Registro de Produccion")
        Exit Sub
    End If
    
    If Libera.ListIndex = 2 Then
        m$ = "Se debe definir si se debe Liberar el area o no"
        A% = MsgBox(m$, 0, "Carga de Registro de Produccion")
        Exit Sub
    End If
    
    If Limpieza.ListIndex = 2 Then
        m$ = "Se debe definir si se debe realizar la limpieza del equipo o no"
        A% = MsgBox(m$, 0, "Carga de Registro de Produccion")
        Exit Sub
    End If
    
    ZZActualizaVersion = "N"
    T$ = "Registro de Produccion"
    m$ = "Desea Actualizar la version"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        ZZActualizaVersion = "S"
        If Trim(ControlCambio.Text) = "" Then
            m$ = "Se debe informar el motivo de cambio de version"
            A% = MsgBox(m$, 0, "Carga de Registro de Produccion")
            Exit Sub
        End If
    End If
    
    
    Version.Text = ""
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and CargaIII.Paso = " + "'" + Paso.Text + "'"
    rsCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        Version.Text = Str$(rstCargaIII!Version)
        rstCargaIII.Close
    End If
    

    Sql1 = "DELETE CargaIII"
    Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
    Sql3 = " and Paso = " + "'" + Paso.Text + "'"
    rsCargaIII = Sql1 + Sql2 + Sql3
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    HastaRenglon = 0
    For iRow = 200 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        Articulo = WVector1.Text
        
        WVector1.Col = 2
        PTerminado = WVector1.Text
        
        WVector1.Col = 3
        Letra = WVector1.Text
        
        WVector1.Col = 4
        XDescripcion = WVector1.Text
        
        WVector1.Col = 5
        XCantidad = WVector1.Text
            
        If Articulo <> "" Or PTerminado <> "" Or Letra <> "" Or XDescripcion <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    

    WRenglon = 0
    WBlanco = ""
    WItem = 0
    Rem WSuma = 0
    
    ZZImprePeso = ""
    ZZImpreHumedad = ""
    
    If Val(Equipo.Text) <> 0 Then
    
        WItem = 1
        
        DesEquipo.Caption = ""
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + Equipo.Text + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            ZZPoe = IIf(IsNull(rstEquipo!Poe), "", rstEquipo!Poe)
            ZZIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            ZZPoeLimpieza = IIf(IsNull(rstEquipo!PoeLimpieza), "", rstEquipo!PoeLimpieza)
            rstEquipo.Close
        End If
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "G"
        XDescripcion = " Utilizar Equipo : " + ZZIdentificacion
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = Trim(Str$(Val(Paso.Text))) + "." + Trim(Str$(WItem))
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        Rem BY NAN
        ZSql = ZSql + "'" + Left$(XDescripcion, 70) + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem WSuma = WSuma + 1
        Rem If WSuma = 55 Then
        Rem     Call Salto
        Rem     WSuma = 0
        Rem End
        
        
        
        
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "G"
        XDescripcion = " Operar Equipo segun POE " + ZZPoe
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "G"
        XDescripcion = ""
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    
    
    If Libera.ListIndex = 1 Then
    
        WItem = WItem + 1
        
        Articulo = ""
        PTerminado = ""
        Letra = ""
        XDescripcion = " Se debe liberar el área / equipo"
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = Trim(Str$(Val(Paso.Text))) + "." + Trim(Str$(WItem))
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        Articulo = ""
        PTerminado = ""
        Letra = ""
        XDescripcion = ""
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    For iRow = 1 To HastaRenglon
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        Articulo = WVector1.Text
        
        WVector1.Col = 2
        PTerminado = WVector1.Text
        
        WVector1.Col = 3
        Letra = WVector1.Text
        
        WVector1.Col = 4
        XDescripcion = WVector1.Text
        
        WVector1.Col = 5
        XCantidad = WVector1.Text
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        If WBlanco = "" Then
            WItem = WItem + 1
            WImpreItem = "S"
        End If
        
        ZZTipo = "S"
        If WImpreItem = "S" Then
            ZZItem = Trim(Str$(Val(Paso.Text))) + "." + Trim(Str$(WItem))
            WImpreItem = "N"
                Else
            ZZItem = ""
        End If
        
        If Peso.ListIndex = 1 And iRow = 1 Then
            ZZImprePeso = "S"
                Else
            ZZImprePeso = ""
        End If
        
        If Humedad.ListIndex = 1 And iRow = 1 Then
            ZZImpreHumedad = "S"
                Else
            ZZImpreHumedad = ""
        End If
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        ZZImprePeso = ""
        ZZImpreHumedad = ""
        If Trim(XDescripcion) = "" Then
            WBlanco = ""
                Else
            WBlanco = "N"
        End If
        
        If Trim(XDescripcion) = "SOLICITAR CONTROL INTERMEDIO a LAB CC" Or UCase(Trim(XDescripcion)) = "REGISTRO DE RESULTADOS" Then
        
            For IRowII = 1 To 100
        
                DescriI = Trim(WVector2.TextMatrix(IRowII, 2))
                DescriII = Trim(WVector2.TextMatrix(IRowII, 3))
                
                If DescriI <> "" Then
                
                    Articulo = ""
                    PTerminado = ""
                    Letra = ""
                    Rem XDescripcion = DescriI + ":" + DescriII
                    XDescripcion = "Metodo " + Trim(WVector2.TextMatrix(IRowII, 1)) + " : " + DescriII
                    XCantidad = ""
            
                    WRenglon = WRenglon + 1
                    Auxi = Str$(WRenglon)
                    Call Ceros(Auxi, 2)
        
                    XPaso = Paso.Text
                    Call Ceros(XPaso, 4)
                        
                    WClave = Terminado.Text + XPaso + Auxi
        
                    ZZTipo = "N"
                    ZZItem = ""
            
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO CargaIII ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Terminado ,"
                    ZSql = ZSql + "Paso ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Articulo ,"
                    ZSql = ZSql + "PTerminado ,"
                    ZSql = ZSql + "Letra ,"
                    ZSql = ZSql + "Descripcion ,"
                    ZSql = ZSql + "Cantidad ,"
                    ZSql = ZSql + "Equipo ,"
                    ZSql = ZSql + "Peso ,"
                    ZSql = ZSql + "Tipo ,"
                    ZSql = ZSql + "Item ,"
                    ZSql = ZSql + "Epp ,"
                    ZSql = ZSql + "DesEpp ,"
                    ZSql = ZSql + "CorteItem ,"
                    ZSql = ZSql + "ImprePeso ,"
                    ZSql = ZSql + "Libera ,"
                    ZSql = ZSql + "Limpieza ,"
                    ZSql = ZSql + "Metodo ,"
                    ZSql = ZSql + "Humedad ,"
                    ZSql = ZSql + "ImpreHumedad )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WClave + "',"
                    ZSql = ZSql + "'" + Terminado.Text + "',"
                    ZSql = ZSql + "'" + Paso.Text + "',"
                    ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                    ZSql = ZSql + "'" + Articulo + "',"
                    ZSql = ZSql + "'" + PTerminado + "',"
                    ZSql = ZSql + "'" + Letra + "',"
                    ZSql = ZSql + "'" + Left$(XDescripcion, 70) + "',"
                    ZSql = ZSql + "'" + XCantidad + "',"
                    ZSql = ZSql + "'" + Equipo.Text + "',"
                    ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
                    ZSql = ZSql + "'" + ZZTipo + "',"
                    ZSql = ZSql + "'" + ZZItem + "',"
                    ZSql = ZSql + "'" + Epp.Text + "',"
                    ZSql = ZSql + "'" + Left$(DesEpp.Caption, 70) + "',"
                    ZSql = ZSql + "'" + Str$(WItem) + "',"
                    ZSql = ZSql + "'" + ZZImprePeso + "',"
                    ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
                    ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
                    ZSql = ZSql + "'" + Metodo.Text + "',"
                    ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
                    ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
                    rsCargaIII = ZSql
                    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            Next IRowII
            
        End If
            
    Next iRow
    
    If Peso.ListIndex = 1 Then
    
        WItem = WItem + 1
        
        Articulo = ""
        PTerminado = ""
        Letra = ""
        XDescripcion = " Registrar tipo de envase y peso neto en tabla : "
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = Trim(Str$(Val(Paso.Text))) + "." + Trim(Str$(WItem))
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "X"
        XDescripcion = " REGISTRO DE PESOS POR ENVASE"
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        Articulo = ""
        PTerminado = ""
        Letra = ""
        XDescripcion = ""
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    
    
        
        
    
    If Limpieza.ListIndex = 1 Then
    
        WItem = WItem + 1
        
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + Equipo.Text + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            ZZPoe = IIf(IsNull(rstEquipo!Poe), "", rstEquipo!Poe)
            ZZIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            ZZPoeLimpieza = IIf(IsNull(rstEquipo!PoeLimpieza), "", rstEquipo!PoeLimpieza)
            rstEquipo.Close
        End If
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "G"
        XDescripcion = " Se debe realizar la limpieza del equipo segun POE : " + Trim(ZZPoeLimpieza) + " Metodo:" + Metodo.Text
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = Trim(Str$(Val(Paso.Text))) + "." + Trim(Str$(WItem))
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        Rem BY NAN
        ZSql = ZSql + "'" + Left$(XDescripcion, 90) + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
        
        
        Articulo = ""
        PTerminado = ""
        Letra = "G"
        XDescripcion = " Registrar en Bitacora"
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        Rem BY NAN
        ZSql = ZSql + "'" + Left$(XDescripcion, 90) + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
        
        
        
        Articulo = ""
        PTerminado = ""
        Letra = ""
        XDescripcion = ""
        XCantidad = ""
            
        WRenglon = WRenglon + 1
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 2)
        
        XPaso = Paso.Text
        Call Ceros(XPaso, 4)
                        
        WClave = Terminado.Text + XPaso + Auxi
        
        ZZTipo = "N"
        ZZItem = ""
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CargaIII ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Paso ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "PTerminado ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Equipo ,"
        ZSql = ZSql + "Peso ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Epp ,"
        ZSql = ZSql + "DesEpp ,"
        ZSql = ZSql + "CorteItem ,"
        ZSql = ZSql + "ImprePeso ,"
        ZSql = ZSql + "Libera ,"
        ZSql = ZSql + "Limpieza ,"
        ZSql = ZSql + "Metodo ,"
        ZSql = ZSql + "Humedad ,"
        ZSql = ZSql + "ImpreHumedad )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + Terminado.Text + "',"
        ZSql = ZSql + "'" + Paso.Text + "',"
        ZSql = ZSql + "'" + Str$(WRenglon) + "',"
        ZSql = ZSql + "'" + Articulo + "',"
        ZSql = ZSql + "'" + PTerminado + "',"
        ZSql = ZSql + "'" + Letra + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XCantidad + "',"
        ZSql = ZSql + "'" + Equipo.Text + "',"
        ZSql = ZSql + "'" + Str$(Peso.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + Epp.Text + "',"
        ZSql = ZSql + "'" + Left$(DesEpp.Caption, 50) + "',"
        ZSql = ZSql + "'" + Str$(WItem) + "',"
        ZSql = ZSql + "'" + ZZImprePeso + "',"
        ZSql = ZSql + "'" + Str$(Libera.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Limpieza.ListIndex) + "',"
        ZSql = ZSql + "'" + Metodo.Text + "',"
        ZSql = ZSql + "'" + Str$(Humedad.ListIndex) + "',"
        ZSql = ZSql + "'" + ZZImpreHumedad + "')"
            
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " ControlCambio = " + "'" + ControlCambio.Text + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    ZZVersion = Version.Text
    ZZFechaVersionII = FechaVersion.Text
    ZZOrdFechaVersionII = Right$(ZZFechaVersionII, 4) + Mid$(ZZFechaVersionII, 4, 2) + Left$(ZZFechaVersionII, 2)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " Version = " + "'" + ZZVersion + "',"
    ZSql = ZSql + " FechaVersion = " + "'" + ZZFechaVersionII + "',"
    ZSql = ZSql + " OrdFechaVersion = " + "'" + ZZOrdFechaVersionII + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    If ZZActualizaVersion = "S" Then
    
        ZZFechaVersionII = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZZOrdFechaVersionII = Right$(ZZFechaVersionII, 4) + Mid$(ZZFechaVersionII, 4, 2) + Left$(ZZFechaVersionII, 2)
        
        ZZVersion = Str$(Val(Version.Text) + 1)
        Call Ceros(ZZVersion, 4)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CargaIII SET "
        ZSql = ZSql + " Version = " + "'" + ZZVersion + "',"
        ZSql = ZSql + " FechaVersion = " + "'" + ZZFechaVersionII + "',"
        ZSql = ZSql + " OrdFechaVersion = " + "'" + ZZOrdFechaVersionII + "'"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
        spCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
        Erase ZZPasa
        ZLugarPasa = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaIII"
        ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " Order by Clave"
        spCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIII.RecordCount > 0 Then
        
            With rstCargaIII
            
                .MoveFirst
                If .NoMatch = False Then
                    Do
                    
                        ZLugarPasa = ZLugarPasa + 1
                        
                        ZZPasa(ZLugarPasa, 1) = rstCargaIII!Terminado
                        ZZPasa(ZLugarPasa, 2) = rstCargaIII!Paso
                        ZZPasa(ZLugarPasa, 3) = rstCargaIII!Renglon
                        ZZPasa(ZLugarPasa, 4) = rstCargaIII!Articulo
                        ZZPasa(ZLugarPasa, 5) = rstCargaIII!PTerminado
                        ZZPasa(ZLugarPasa, 6) = rstCargaIII!Letra
                        ZZPasa(ZLugarPasa, 7) = rstCargaIII!Descripcion
                        ZZPasa(ZLugarPasa, 8) = Str$(rstCargaIII!Cantidad)
                        ZZPasa(ZLugarPasa, 10) = IIf(IsNull(rstCargaIII!Partida), "", rstCargaIII!Partida)
                        ZZPasa(ZLugarPasa, 11) = IIf(IsNull(rstCargaIII!CantidadPartida), "", rstCargaIII!CantidadPartida)
                        ZZPasa(ZLugarPasa, 12) = IIf(IsNull(rstCargaIII!Equipo), "", rstCargaIII!Equipo)
                        ZZPasa(ZLugarPasa, 13) = IIf(IsNull(rstCargaIII!Peso), "0", rstCargaIII!Peso)
                        ZZPasa(ZLugarPasa, 14) = IIf(IsNull(rstCargaIII!Tipo), "", rstCargaIII!Tipo)
                        ZZPasa(ZLugarPasa, 15) = IIf(IsNull(rstCargaIII!Item), "", rstCargaIII!Item)
                        ZZPasa(ZLugarPasa, 16) = IIf(IsNull(rstCargaIII!Epp), "", rstCargaIII!Epp)
                        ZZPasa(ZLugarPasa, 17) = IIf(IsNull(rstCargaIII!DesEpp), "", rstCargaIII!DesEpp)
                        ZZPasa(ZLugarPasa, 18) = IIf(IsNull(rstCargaIII!CorteItem), "", rstCargaIII!CorteItem)
                        ZZPasa(ZLugarPasa, 19) = IIf(IsNull(rstCargaIII!ImprePeso), "", rstCargaIII!ImprePeso)
                        ZZPasa(ZLugarPasa, 20) = IIf(IsNull(rstCargaIII!Humedad), "0", rstCargaIII!Humedad)
                        ZZPasa(ZLugarPasa, 21) = IIf(IsNull(rstCargaIII!ImpreHumedad), "0", rstCargaIII!ImpreHumedad)
                        ZZPasa(ZLugarPasa, 22) = IIf(IsNull(rstCargaIII!Libera), "0", rstCargaIII!Libera)
                        ZZPasa(ZLugarPasa, 23) = IIf(IsNull(rstCargaIII!Limpieza), "0", rstCargaIII!Limpieza)
                        ZZPasa(ZLugarPasa, 24) = IIf(IsNull(rstCargaIII!Metodo), "0", rstCargaIII!Metodo)
                        ZZPasa(ZLugarPasa, 25) = IIf(IsNull(rstCargaIII!ControlCambio), "", rstCargaIII!ControlCambio)
                        ZZPasa(ZLugarPasa, 26) = IIf(IsNull(rstCargaIII!Version), "", rstCargaIII!Version)
                        ZZPasa(ZLugarPasa, 27) = rstCargaIII!Clave
                        
                        .MoveNext
                        
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                    Loop
                End If
                
            End With
            rstCargaIII.Close
        
        End If
        
        For Ciclopasa = 1 To ZLugarPasa
            
            ZZTerminado = ZZPasa(Ciclopasa, 1)
            ZZPaso = ZZPasa(Ciclopasa, 2)
            ZZRenglon = ZZPasa(Ciclopasa, 3)
            ZZArticulo = ZZPasa(Ciclopasa, 4)
            ZZPTerminado = ZZPasa(Ciclopasa, 5)
            ZZLetra = ZZPasa(Ciclopasa, 6)
            ZZDescripcion = ZZPasa(Ciclopasa, 7)
            ZZCantidad = ZZPasa(Ciclopasa, 8)
            ZZCantidadII = ZZPasa(Ciclopasa, 9)
            ZZPartida = ZZPasa(Ciclopasa, 10)
            ZZCantidadPartida = ZZPasa(Ciclopasa, 11)
            ZZEquipo = ZZPasa(Ciclopasa, 12)
            ZZPeso = ZZPasa(Ciclopasa, 13)
            ZZTipo = ZZPasa(Ciclopasa, 14)
            ZZItem = ZZPasa(Ciclopasa, 15)
            ZZEpp = ZZPasa(Ciclopasa, 16)
            ZZDesEpp = ZZPasa(Ciclopasa, 17)
            ZZCorteItem = ZZPasa(Ciclopasa, 18)
            ZZImprePeso = ZZPasa(Ciclopasa, 19)
            ZZHumedad = ZZPasa(Ciclopasa, 20)
            ZZImpreHumedad = ZZPasa(Ciclopasa, 21)
            ZZLibera = ZZPasa(Ciclopasa, 22)
            ZZLimpieza = ZZPasa(Ciclopasa, 23)
            ZZMetodo = ZZPasa(Ciclopasa, 24)
            ZZControlCambio = ZZPasa(Ciclopasa, 25)
            ZZVersion = ZZPasa(Ciclopasa, 26)
            Call Ceros(ZZVersion, 4)
            ZZClave = Left$(ZZPasa(Ciclopasa, 27), 12) + ZZVersion + Mid$(ZZPasa(Ciclopasa, 27), 13, 10)
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaIIIVersion ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Paso ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "PTerminado ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Equipo ,"
            ZSql = ZSql + "Peso ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Epp ,"
            ZSql = ZSql + "DesEpp ,"
            ZSql = ZSql + "CorteItem ,"
            ZSql = ZSql + "ImprePeso ,"
            ZSql = ZSql + "Libera ,"
            ZSql = ZSql + "Limpieza ,"
            ZSql = ZSql + "Metodo ,"
            ZSql = ZSql + "Humedad ,"
            ZSql = ZSql + "ImpreHumedad ,"
            ZSql = ZSql + "ControlCambio ,"
            ZSql = ZSql + "FechaVersionI ,"
            ZSql = ZSql + "OrdFechaVersionI ,"
            ZSql = ZSql + "FechaVersionII ,"
            ZSql = ZSql + "OrdFechaVersionII )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZTerminado + "',"
            ZSql = ZSql + "'" + ZZVersion + "',"
            ZSql = ZSql + "'" + ZZPaso + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + Trim(ZZArticulo) + "',"
            ZSql = ZSql + "'" + Trim(ZZPTerminado) + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + Left$(ZZDescripcion, 70) + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZEquipo + "',"
            ZSql = ZSql + "'" + ZZPeso + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZItem + "',"
            ZSql = ZSql + "'" + ZZEpp + "',"
            ZSql = ZSql + "'" + Left$(ZZDesEpp, 70) + "',"
            ZSql = ZSql + "'" + ZZItem + "',"
            ZSql = ZSql + "'" + ZZImprePeso + "',"
            ZSql = ZSql + "'" + ZZLibera + "',"
            ZSql = ZSql + "'" + ZZLimpieza + "',"
            ZSql = ZSql + "'" + ZZMetodo + "',"
            ZSql = ZSql + "'" + ZZHumedad + "',"
            ZSql = ZSql + "'" + ZZImpreHumedad + "',"
            ZSql = ZSql + "'" + ZZControlCambio + "',"
            ZSql = ZSql + "'" + ZZFechaVersionI + "',"
            ZSql = ZSql + "'" + ZZOrdFechaVersionI + "',"
            ZSql = ZSql + "'" + ZZFechaVersionII + "',"
            ZSql = ZSql + "'" + ZZOrdFechaVersionII + "')"
                
            rsCargaIIIVersion = ZSql
            Set rstCargaIIIVersion = db.OpenRecordset(rsCargaIIIVersion, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Ciclopasa
    
    End If
    
    Call Limpia_Click

    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Tablas.Tab = 0
        
    Terminado.SetFocus
        
End Sub

Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    
    Tablas.Tab = 0

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    Equipo.Text = ""
    DesEquipo.Caption = ""
    Epp.Text = ""
    DesEpp.Caption = ""
    Metodo.Text = ""
    Version.Text = ""
    Version.Locked = True
    FechaVersion.Text = "  /  /    "
    ControlCambio.Text = ""
    
    Peso.ListIndex = 2
    Humedad.ListIndex = 2
    Libera.ListIndex = 2
    Limpieza.ListIndex = 2
    
    Renglon = 0
    Graba.Enabled = True
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    Terminado.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Terminado.Text = WIndice.List(Indice)
            Call Terminado_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = Trim(rstArticulo!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 2
            Indice = Pantalla.ListIndex
            WPTerminado = WIndice.List(Indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WPTerminado + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = Trim(rstTerminado!Codigo)
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 3
                rstTerminado.Close
                Call StartEdit
            End If
            Rem Ayuda.Visible = False
            
        Case 4
            Indice = Pantalla.ListIndex
            Equipo.Text = WIndice.List(Indice)
            Call Equipo_KeyPress(13)
            
        Case 5
            Indice = Pantalla.ListIndex
            Epp.Text = WIndice.List(Indice)
            Call Epp_KeyPress(13)
            
        Case 6
            Indice = Pantalla.ListIndex
            WVector1.Col = 4
            WVector1.Text = WIndice.List(Indice)
            Call StartEdit
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    Peso.Clear
    
    Peso.AddItem "No Solicita Peso"
    Peso.AddItem "Solicita Peso"
    Peso.AddItem ""
    
    Peso.ListIndex = 2
    
    Humedad.Clear
    
    Humedad.AddItem "No Controla Humedad"
    Humedad.AddItem "Controla Humedad"
    Humedad.AddItem ""
    
    Humedad.ListIndex = 2
    
    Libera.Clear
    
    Libera.AddItem "No se debe liberar el area"
    Libera.AddItem "Se debe liberar el area"
    Libera.AddItem ""
    
    Libera.ListIndex = 2
    
    Limpieza.Clear
    
    Limpieza.AddItem "No se debe limpiar el equipo"
    Limpieza.AddItem "Se debe limpiar el equipo"
    Limpieza.AddItem ""
    
    Limpieza.ListIndex = 2

    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    Equipo.Text = ""
    DesEquipo.Caption = ""
    Epp.Text = ""
    DesEpp.Caption = ""
    Metodo.Text = ""
    Version.Text = ""
    Version.Locked = True
    FechaVersion.Text = "  /  /    "
    ControlCambio.Text = ""
    
    Renglon = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WRenglon = 0
    
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and CargaIII.Paso = " + "'" + Paso.Text + "'"
    ZSql = ZSql + " and CargaIII.Tipo <> 'N'"
    ZSql = ZSql + " Order by CargaIII.Terminado, CargaIII.Paso, CargaIII.Renglon"
    
    rsCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        With rstCargaIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    Equipo.Text = IIf(IsNull(rstCargaIII!Equipo), "", rstCargaIII!Equipo)
                    Peso.ListIndex = IIf(IsNull(rstCargaIII!Peso), "0", rstCargaIII!Peso)
                    Humedad.ListIndex = IIf(IsNull(rstCargaIII!Humedad), "0", rstCargaIII!Humedad)
                    Epp.Text = IIf(IsNull(rstCargaIII!Epp), "", rstCargaIII!Epp)
                    Libera.ListIndex = IIf(IsNull(rstCargaIII!Libera), "0", rstCargaIII!Libera)
                    Limpieza.ListIndex = IIf(IsNull(rstCargaIII!Limpieza), "0", rstCargaIII!Limpieza)
                    Metodo.Text = IIf(IsNull(rstCargaIII!Metodo), "0", rstCargaIII!Metodo)
                    Version.Text = IIf(IsNull(rstCargaIII!Version), "", rstCargaIII!Version)
                    FechaVersion.Text = IIf(IsNull(rstCargaIII!FechaVersion), "  /  /    ", rstCargaIII!FechaVersion)
                    ControlCambio.Text = IIf(IsNull(rstCargaIII!ControlCambio), "", rstCargaIII!ControlCambio)
                    
                    If Val(Version.Text) = 0 Then
                        Version.Locked = False
                            Else
                        Version.Locked = True
                    End If
                    
                    WVector1.Col = 1
                    ZZZArticulo = IIf(IsNull(rstCargaIII!Articulo), "", rstCargaIII!Articulo)
                    WVector1.Text = Trim(ZZZArticulo)
            
                    WVector1.Col = 2
                    ZZZPTerminado = IIf(IsNull(rstCargaIII!PTerminado), "", rstCargaIII!PTerminado)
                    WVector1.Text = Trim(ZZZPTerminado)
            
                    WVector1.Col = 3
                    ZZZLetra = IIf(IsNull(rstCargaIII!Letra), "", rstCargaIII!Letra)
                    WVector1.Text = Trim(ZZZLetra)
            
                    WVector1.Col = 4
                    ZZZDescripcion = IIf(IsNull(rstCargaIII!Descripcion), "", rstCargaIII!Descripcion)
                    WVector1.Text = Trim(ZZZDescripcion)
            
                    WVector1.Col = 5
                    If rstCargaIII!Cantidad <> 0 Then
                        ZZZCantidad = IIf(IsNull(rstCargaIII!Cantidad), "0", rstCargaIII!Cantidad)
                        WVector1.Text = Str$(ZZZCantidad)
                        WVector1.Text = Pusing("###.#####", WVector1.Text)
                            Else
                        WVector1.Text = ""
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIII.Close
    End If
    
    
    
    WRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaV"
    ZSql = ZSql + " Where CargaV.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and CargaV.Paso = " + "'" + Paso.Text + "'"
    ZSql = ZSql + " Order by CargaV.Clave"
    
    rsCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaV.RecordCount > 0 Then
        With rstCargaV
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    
                    WVector2.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector2.Col = 1
                    ZZZEnsayo = IIf(IsNull(rstCargaV!Ensayo), "0", rstCargaV!Ensayo)
                    WVector2.Text = Trim(Str$(ZZZEnsayo))
            
                    WVector2.Col = 2
                    WVector2.Text = ""
            
                    WVector2.Col = 3
                    ZZZValor = IIf(IsNull(rstCargaV!Valor), "", rstCargaV!Valor)
                    WVector2.Text = Trim(ZZZValor)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaV.Close
    End If
    
    
    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 9
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
    
    For Ciclo = 1 To WRenglon
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ensayos"
        ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.TextMatrix(Ciclo, 1) + "'"
        spEnsayos = ZSql
        Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayos.RecordCount > 0 Then
            WVector2.TextMatrix(Ciclo, 2) = Trim(rstEnsayos!Descripcion)
            rstEnsayos.Close
        End If
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    Sql1 = "Select *"
    Sql2 = " FROM Terminado"
    Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
    spTerminado = Sql1 + Sql2 + Sql3
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = Trim(rstTerminado!Descripcion)
        rstTerminado.Close
    End If
    
    DesEquipo.Caption = ""
    Sql1 = "Select *"
    Sql2 = " FROM Equipo"
    Sql3 = " Where Equipo.Codigo = " + "'" + Equipo.Text + "'"
    spEquipo = Sql1 + Sql2 + Sql3
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        DesEquipo.Caption = Trim(rstEquipo!Descripcion)
        rstEquipo.Close
    End If
    
    DesEpp.Caption = ""
    Sql1 = "Select *"
    Sql2 = " FROM MaterialAuxiliar"
    Sql3 = " Where MaterialAuxiliar.Codigo = " + "'" + Epp.Text + "'"
    spMaterialAuxiliar = Sql1 + Sql2 + Sql3
    Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMaterialAuxiliar.RecordCount > 0 Then
        DesEpp.Caption = Trim(rstMaterialAuxiliar!Descripcion)
        rstMaterialAuxiliar.Close
    End If
    
    Tablas.Tab = 0
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    Call StartEdit
    
    Graba.Enabled = True

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Terminado.Text = UCase(Terminado.Text)
    
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = Trim(rstTerminado!Descripcion)
            rstTerminado.Close
            Paso.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Terminado.Text = "  -     -   "
        DesTerminado.Caption = ""
    End If
End Sub

Private Sub Paso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Existe = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaIII"
        ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and CargaIII.Paso = " + "'" + Paso.Text + "'"
        rsCargaIII = ZSql
        Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaIII.RecordCount > 0 Then
            rstCargaIII.Close
            Existe = "S"
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaV"
        ZSql = ZSql + " Where CargaV.Terminado = " + "'" + Terminado.Text + "'"
        ZSql = ZSql + " and CargaV.Paso = " + "'" + Paso.Text + "'"
        rsCargaV = ZSql
        Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaV.RecordCount > 0 Then
            rstCargaV.Close
            Existe = "S"
        End If
        
        If Existe = "S" Then
            Call Proceso_Click
                Else
            Graba.Enabled = True
            WTerminado = Terminado.Text
            WPaso = Terminado.Text
            Terminado.Text = WTerminado
            Paso.Text = Paso
            Call Limpia_Vector
            Call Limpia_VectorII
            Tablas.Tab = 0
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            WVector2.TopRow = 1
            WVector2.Col = 1
            WVector2.Row = 1
            Rem Call StartEdit
            Peso.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Paso.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0, 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTerminado!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
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
                rstTerminado.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM MaterialAuxiliar"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spMaterialAuxiliar = Sql1 + Sql2 + Sql3 + Sql4
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                With rstMaterialAuxiliar
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMaterialAuxiliar.Close
            End If
            
        Case 3
            XEmpresa = Wempresa
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 9
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
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                With rstEnsayos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayos.Close
            End If
            
            Call Conecta_Empresa
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Equipo"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spEquipo = ZSql
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                With rstEquipo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEquipo.Close
            End If
            
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM MaterialAuxiliar"
            ZSql = ZSql + " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Codigo"
            spMaterialAuxiliar = ZSql
            Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMaterialAuxiliar.RecordCount > 0 Then
                With rstMaterialAuxiliar
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMaterialAuxiliar.Close
            End If
            
        Case 6
            Sql1 = "Select *"
            Sql2 = " FROM TextoFijo"
            Sql3 = " Where Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            Sql4 = " Order by Codigo"
            spTextoFijo = Sql1 + Sql2 + Sql3 + Sql4
            Set rstTextoFijo = db.OpenRecordset(spTextoFijo, dbOpenSnapshot, dbSQLPassThrough)
            If rstTextoFijo.RecordCount > 0 Then
                With rstTextoFijo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Descripcion
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTextoFijo.Close
            End If
            
        Case Else
    End Select
            
    End If

End Sub

Private Sub Terminado_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Productos Terminados"
    Opcion.AddItem "Material Auxiliar a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

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
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
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
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
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
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
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
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM articulo"
            Sql3 = " Where articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstArticulo!Descripcion)
                End If
                WVector1.Col = 3
                rstArticulo.Close
            End If
            
        Case 2
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + WVector1.Text + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WVector1.Col = 4
                If WVector1.Text = "" Then
                    WVector1.Text = Trim(rstTerminado!Descripcion)
                End If
                WVector1.Col = 3
                rstTerminado.Close
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    RenglonAuxiliar = WVector1.Row

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 200 To 1 Step -1
        
        WVector1.Row = iRow
            
        WVector1.Col = 1
        Articulo = WVector1.Text
        
        WVector1.Col = 2
        PTerminado = WVector1.Text
        
        WVector1.Col = 3
        Letra = WVector1.Text
        
        WVector1.Col = 4
        XDescripcion = WVector1.Text
        
        WVector1.Col = 5
        XCantidad = WVector1.Text
            
        If Articulo <> "" Or PTerminado <> "" Or Letra <> "" Or XDescripcion <> "" Or XCantidad <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        If Ciclo <> RenglonAuxiliar Then
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
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 4 Then
    
        Opcion.Clear
        
        Opcion.AddItem "Productos Terminados"
        Opcion.AddItem "Materia Prima a Utilizar"
        Opcion.AddItem "Productos Terminados a Utilizar"
        Opcion.AddItem "Ensayos"
        Opcion.AddItem "Equipos"
        Opcion.AddItem "Epp"
        Opcion.AddItem "Texto Fijo"
    
        Opcion.ListIndex = 6
    
    End If
    
End Sub



Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
    End If
    
    If WVector1.Col = 2 Then

    Opcion.Clear
    
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click
    
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
    WVector1.Cols = 6
    WVector1.FixedRows = 1
    WVector1.Rows = 201
    
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
    
    WVector1.ColWidth(0) = 300
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "M.Prima"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "P.Terminado"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 6000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 70
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###.#####"
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

    For Ciclo = 1 To 199
        WVector1.TextMatrix(Ciclo, 0) = Trim(Str$(Ciclo))
    Next Ciclo


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




Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 4
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
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
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Ensayos"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 4
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 4500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Valor"
                WVector2.ColWidth(Ciclo) = 5000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector2.Text
        Rem WTitulo(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        Rem WTitulo(Ciclo).Top = WVector2.CellTop + WVector2.Top
        Rem WTitulo(Ciclo).Width = WVector2.CellWidth
        Rem WTitulo(Ciclo).Height = WVector2.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 0
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case 1
            WVector2.Col = 1
            WVector2.Row = 1
        Case Else
    End Select
End Sub

Private Sub Conecta_Empresa()

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
        Case Else
    End Select

End Sub

Private Sub AltaPoe_Click()

    File1.Pattern = "*.DOC"
    File1.ListIndex = 0
    
    Rem IngresaFotos.Height = 4575
    Rem IngresaFotos.Left = 1080
    Rem IngresaFotos.Top = 120
    Rem IngresaFotos.Width = 6500
    
    IngresaPoe.Height = 4815
    IngresaPoe.Left = 600
    IngresaPoe.Top = 840
    IngresaPoe.Width = 5500
    
    IngresaPoe.Visible = True
    
End Sub

Private Sub File1_dblClick()

    ZLargo = Len(File1.filename)

    If WTexto1.Visible = True Then
        WVector1.TextMatrix(WVector1.Row, WVector1.Col) = WTexto1.Text + " " + Left$(File1.filename, ZLargo - 4)
    End If
    
    If WTexto2.Visible = True Then
        WVector1.TextMatrix(WVector1.Row, WVector1.Col) = WTexto2.Text + " " + Left$(File1.filename, ZLargo - 4)
    End If
    
    Call StartEdit
    IngresaPoe.Visible = False

End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive    ' Establece la ruta del directorio.
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Establece la ruta del archivo.
End Sub




Private Sub GrabaRegistro_Click()

    TerminadoII.Text = "  -     -   "
    DesTerminadoII.Caption = ""
    PasoII.Text = ""
    VersionII.Text = ""
    
    PantaGraba.Visible = True
    
    TerminadoII.SetFocus

End Sub

Private Sub TerminadoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        TerminadoII.Text = UCase(TerminadoII.Text)
    
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + TerminadoII.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminadoII.Caption = Trim(rstTerminado!Descripcion)
            rstTerminado.Close
            PasoII.SetFocus
                Else
            TerminadoII.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        TerminadoII.Text = "  -     -   "
        DesTerminadoII.Caption = ""
    End If
End Sub

Private Sub PasoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VersionII.SetFocus
    End If
    If KeyAscii = 27 Then
        PasoII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub VersionII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TerminadoII.SetFocus
    End If
    If KeyAscii = 27 Then
        PasoII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CierraII_Click()
    PantaGraba.Visible = False
End Sub

Private Sub GrabaII_Click()

    Call Limpia_Vector
    
    WRenglon = 0
    
    ZSql = " "
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIIIVersion"
    ZSql = ZSql + " Where CargaIIIVersion.Terminado = " + "'" + TerminadoII.Text + "'"
    ZSql = ZSql + " and CargaIIIVersion.Paso = " + "'" + PasoII.Text + "'"
    ZSql = ZSql + " and CargaIIIVersion.Version = " + "'" + VersionII.Text + "'"
    ZSql = ZSql + " and CargaIIIVersion.Tipo <> 'N'"
    ZSql = ZSql + " Order by CargaIIIVersion.Terminado, CargaIIIVersion.Paso, CargaIIIVersion.Renglon"
    
    rsCargaIIIVersion = ZSql
    Set rstCargaIIIVersion = db.OpenRecordset(rsCargaIIIVersion, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIIIVersion.RecordCount > 0 Then
        With rstCargaIIIVersion
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    Equipo.Text = IIf(IsNull(rstCargaIIIVersion!Equipo), "", rstCargaIIIVersion!Equipo)
                    Peso.ListIndex = IIf(IsNull(rstCargaIIIVersion!Peso), "0", rstCargaIIIVersion!Peso)
                    Humedad.ListIndex = IIf(IsNull(rstCargaIIIVersion!Humedad), "0", rstCargaIIIVersion!Humedad)
                    Epp.Text = IIf(IsNull(rstCargaIIIVersion!Epp), "", rstCargaIIIVersion!Epp)
                    Libera.ListIndex = IIf(IsNull(rstCargaIIIVersion!Libera), "0", rstCargaIIIVersion!Libera)
                    Limpieza.ListIndex = IIf(IsNull(rstCargaIIIVersion!Limpieza), "0", rstCargaIIIVersion!Limpieza)
                    Metodo.Text = IIf(IsNull(rstCargaIIIVersion!Metodo), "0", rstCargaIIIVersion!Metodo)
                    Rem Version.Text = IIf(IsNull(rstCargaIIIVersion!Version), "", rstCargaIIIVersion!Version)
                    Rem FechaVersion.Text = IIf(IsNull(rstCargaIIIVersion!FechaVersion), "  /  /    ", rstCargaIIIVersion!FechaVersion)
                    Rem ControlCambio.Text = IIf(IsNull(rstCargaIIIVersion!ControlCambio), "", rstCargaIIIVersion!ControlCambio)
                    
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaIIIVersion!Articulo)
            
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaIIIVersion!PTerminado)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCargaIIIVersion!Letra)
            
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCargaIIIVersion!Descripcion)
            
                    WVector1.Col = 5
                    If rstCargaIIIVersion!Cantidad <> 0 Then
                        WVector1.Text = Str$(rstCargaIIIVersion!Cantidad)
                        WVector1.Text = Pusing("###.#####", WVector1.Text)
                            Else
                        WVector1.Text = ""
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIIIVersion.Close
    End If
    
    DesEquipo.Caption = ""
    Sql1 = "Select *"
    Sql2 = " FROM Equipo"
    Sql3 = " Where Equipo.Codigo = " + "'" + Equipo.Text + "'"
    spEquipo = Sql1 + Sql2 + Sql3
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        DesEquipo.Caption = Trim(rstEquipo!Descripcion)
        rstEquipo.Close
    End If
    
    DesEpp.Caption = ""
    Sql1 = "Select *"
    Sql2 = " FROM MaterialAuxiliar"
    Sql3 = " Where MaterialAuxiliar.Codigo = " + "'" + Epp.Text + "'"
    spMaterialAuxiliar = Sql1 + Sql2 + Sql3
    Set rstMaterialAuxiliar = db.OpenRecordset(spMaterialAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMaterialAuxiliar.RecordCount > 0 Then
        DesEpp.Caption = Trim(rstMaterialAuxiliar!Descripcion)
        rstMaterialAuxiliar.Close
    End If
    
    PantaGraba.Visible = False

End Sub

