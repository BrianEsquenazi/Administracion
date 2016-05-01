VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAltaCertificado 
   Caption         =   "Ingreso de Datos a Imprimir en las Certificados de Analisis"
   ClientHeight    =   8190
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   11685
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
      MaxLength       =   50
      TabIndex        =   73
      Text            =   " "
      Top             =   6760
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.CheckBox Opcion10 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   72
      Top             =   6480
      Width           =   255
   End
   Begin VB.CheckBox Opcion9 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox Opcion8 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   70
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox Opcion7 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox Opcion6 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   68
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Opcion5 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox Opcion4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Opcion3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox Opcion2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   64
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Opcion1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   63
      Top             =   1440
      Width           =   255
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   60
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      ItemData        =   "altacertificado.frx":0000
      Left            =   120
      List            =   "altacertificado.frx":0007
      TabIndex        =   58
      Top             =   7080
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   2520
      TabIndex        =   54
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   56
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   55
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label22 
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
         TabIndex        =   57
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.TextBox Valor1010 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   53
      Text            =   " "
      Top             =   6720
      Width           =   4815
   End
   Begin VB.TextBox Valor99 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   52
      Text            =   " "
      Top             =   6240
      Width           =   4815
   End
   Begin VB.TextBox Valor88 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   51
      Text            =   " "
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox Valor77 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   50
      Text            =   " "
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox Valor66 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   49
      Text            =   " "
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Valor55 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   48
      Text            =   " "
      Top             =   3960
      Width           =   4815
   End
   Begin VB.TextBox Valor44 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   47
      Text            =   " "
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox Valor33 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   2760
      Width           =   4815
   End
   Begin VB.TextBox Valor22 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   45
      Text            =   " "
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox Valor11 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   44
      Text            =   " "
      Top             =   1680
      Width           =   4815
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
   Begin Crystal.CrystalReport Lista 
      Left            =   11400
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEspefUnifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
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
      Height          =   780
      Left            =   360
      TabIndex        =   42
      Top             =   7200
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   11160
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox valor10 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   6480
      Width           =   4815
   End
   Begin VB.TextBox valor9 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   39
      Text            =   " "
      Top             =   6000
      Width           =   4815
   End
   Begin VB.TextBox valor8 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   38
      Text            =   " "
      Top             =   5400
      Width           =   4815
   End
   Begin VB.TextBox valor7 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   37
      Text            =   " "
      Top             =   4800
      Width           =   4815
   End
   Begin VB.TextBox valor6 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   36
      Text            =   " "
      Top             =   4320
      Width           =   4815
   End
   Begin VB.TextBox valor5 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   3720
      Width           =   4815
   End
   Begin VB.TextBox valor4 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   3120
      Width           =   4815
   End
   Begin VB.TextBox Valor3 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox valor2 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   32
      Text            =   " "
      Top             =   1965
      Width           =   4815
   End
   Begin VB.TextBox Valor1 
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
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox Ensayo10 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   30
      Text            =   " "
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox Ensayo9 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   29
      Text            =   " "
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Ensayo8 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   28
      Text            =   " "
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Ensayo7 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   27
      Text            =   " "
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Ensayo6 
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
      Height          =   315
      Left            =   720
      MaxLength       =   4
      TabIndex        =   26
      Text            =   " "
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox Ensayo5 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   25
      Text            =   " "
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Ensayo4 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   24
      Text            =   " "
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Ensayo3 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   23
      Text            =   " "
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Ensayo2 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   22
      Text            =   " "
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Ensayo1 
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
      Left            =   720
      MaxLength       =   4
      TabIndex        =   21
      Text            =   " "
      Top             =   1440
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   7200
      Visible         =   0   'False
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
      Height          =   420
      Left            =   9000
      TabIndex        =   5
      Top             =   7560
      Width           =   1095
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
      Height          =   420
      Left            =   7800
      TabIndex        =   4
      Top             =   7560
      Width           =   1095
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
      Left            =   10200
      TabIndex        =   3
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
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
      Left            =   9000
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
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
      Height          =   420
      Left            =   7800
      TabIndex        =   1
      Top             =   7080
      Width           =   1095
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
      TabIndex        =   62
      Top             =   480
      Width           =   1575
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
      Height          =   255
      Left            =   2880
      TabIndex        =   61
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label DesProducto 
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
      Height          =   255
      Left            =   2880
      TabIndex        =   59
      Top             =   120
      Width           =   4980
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
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
      TabIndex        =   43
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Descri10 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   6480
      Width           =   4980
   End
   Begin VB.Label Descri9 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   6000
      Width           =   4980
   End
   Begin VB.Label Descri8 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   5400
      Width           =   4980
   End
   Begin VB.Label Descri7 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   4800
      Width           =   4980
   End
   Begin VB.Label Descri6 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   4320
      Width           =   4980
   End
   Begin VB.Label Descri5 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   3720
      Width           =   4980
   End
   Begin VB.Label Descri4 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   3120
      Width           =   4980
   End
   Begin VB.Label Descri3 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   2520
      Width           =   4980
   End
   Begin VB.Label descri2 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   4980
   End
   Begin VB.Label Descri1 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1440
      Width           =   4980
   End
   Begin VB.Label lblresultado 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valor Standard"
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
      Left            =   6720
      TabIndex        =   10
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lblDescri 
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
      Left            =   1560
      TabIndex        =   9
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lblensayo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ensayo"
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
      Left            =   720
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgAltaCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstAltaCertificado As Recordset
Dim spAltaCertificado As String
Dim XParam As String
Dim EmpresaActual As String
Private WGraba As String
Private WGrabaII As String
Dim CargaEmpresa(12, 2) As String

Private Sub Imprime_Datos()

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
    
    Sql1 = "Select Ensayo1,Ensayo2,Ensayo3,ensayo4,ensayo5,Ensayo6,ensayo7,ensayo8,ensayo9,ensayo10,valor1,valor2,valor3,valor4,valor5,valor6,valor7,valor8,valor9,valor10,valor11,valor22,valor33,valor44,valor55,valor66,valor77,valor88,valor99,valor1010"
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        Ensayo1.Text = rstEspecifUnifica!Ensayo1
        Ensayo2.Text = rstEspecifUnifica!Ensayo2
        Ensayo3.Text = rstEspecifUnifica!Ensayo3
        Ensayo4.Text = rstEspecifUnifica!Ensayo4
        Ensayo5.Text = rstEspecifUnifica!Ensayo5
        Ensayo6.Text = rstEspecifUnifica!Ensayo6
        Ensayo7.Text = rstEspecifUnifica!Ensayo7
        Ensayo8.Text = rstEspecifUnifica!Ensayo8
        Ensayo9.Text = rstEspecifUnifica!Ensayo9
        Ensayo10.Text = rstEspecifUnifica!Ensayo10
        
        Rem ZValor1.Text = rstEspecifUnifica!Valor1
        Rem valor2.Text = rstEspecifUnifica!valor2
        Rem Valor3.Text = rstEspecifUnifica!Valor3
        Rem valor4.Text = rstEspecifUnifica!valor4
        Rem valor5.Text = rstEspecifUnifica!valor5
        Rem valor6.Text = rstEspecifUnifica!valor6
        Rem valor7.Text = rstEspecifUnifica!valor7
        Rem valor8.Text = rstEspecifUnifica!valor8
        Rem valor9.Text = rstEspecifUnifica!valor9
        Rem valor10.Text = rstEspecifUnifica!valor10
        
        ZStd1 = rstEspecifUnifica!Valor1
        ZStd2 = rstEspecifUnifica!valor2
        ZStd3 = rstEspecifUnifica!Valor3
        ZStd4 = rstEspecifUnifica!valor4
        ZStd5 = rstEspecifUnifica!valor5
        ZStd6 = rstEspecifUnifica!valor6
        ZStd7 = rstEspecifUnifica!valor7
        ZStd8 = rstEspecifUnifica!valor8
        ZStd9 = rstEspecifUnifica!valor9
        ZStd10 = rstEspecifUnifica!valor10
        
        Valor11.Text = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
        Valor22.Text = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
        Valor33.Text = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
        Valor44.Text = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
        Valor55.Text = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
        Valor66.Text = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
        Valor77.Text = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
        Valor88.Text = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
        Valor99.Text = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
        Valor1010.Text = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
        
        rstEspecifUnifica.Close
        
    End If
        
            
    Sql1 = "Select desde1,Desde2,Desde3,Desde4,Desde5,Desde6,Desde7,Desde8,desde9,Desde10,Hasta1,Hasta2,Hasta3,Hasta4,Hasta5,Hasta6,Hasta7,Hasta8,Hasta9,Hasta10,Valor1Ing "
    Sql2 = " FROM EspecifUnifica"
    Sql3 = " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecifUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        ZDesde1 = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
        ZDesde2 = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
        ZDesde3 = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
        ZDesde4 = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
        ZDesde5 = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
        ZDesde6 = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
        ZDesde7 = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
        ZDesde8 = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
        ZDesde9 = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
        ZDesde10 = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
        
        ZHasta1 = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
        ZHasta2 = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
        ZHasta3 = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
        ZHasta4 = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
        ZHasta5 = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
        ZHasta6 = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
        ZHasta7 = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
        ZHasta8 = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
        ZHasta9 = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
        ZHasta10 = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
        
        rstEspecifUnifica.Close
        
    End If
    
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        ZDescri1 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
        ZDescri1 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        descri2.Caption = rstEnsayo!Descripcion
        ZDescri2 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        descri2.Caption = ""
        ZDescri2 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        ZDescri3 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
        ZDescri3 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        ZDescri4 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
        ZDescri4 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        ZDescri5 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
        ZDescri5 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        ZDescri6 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
        ZDescri6 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        ZDescri7 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
        ZDescri7 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        ZDescri8 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
        ZDescri8 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        ZDescri9 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
        ZDescri9 = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        ZDescri10 = rstEnsayo!Descripcion
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
        ZDescri10 = ""
    End If
    
    If Val(ZDesde1) <> 0 Or Val(ZZHasta1) <> 0 Then
        Valor1.Text = Trim(ZDesde1) + " - " + Trim(ZHasta1) + " " + Trim(ZDescri1) + " " + Left$(ZStd1, 50)
            Else
        Valor1.Text = Left$(ZStd1, 50)
    End If
    
    If Val(ZDesde2) <> 0 Or Val(ZZHasta2) <> 0 Then
        valor2.Text = Trim(ZDesde2) + " - " + Trim(ZHasta2) + " " + Trim(ZDescri2) + " " + Left$(ZStd2, 50)
            Else
        valor2.Text = Left$(ZStd2, 50)
    End If
    
    If Val(ZDesde3) <> 0 Or Val(ZZHasta3) <> 0 Then
        Valor3.Text = Trim(ZDesde3) + " - " + Trim(ZHasta3) + " " + Trim(ZDescri3) + " " + Left$(ZStd3, 50)
            Else
        Valor3.Text = Left$(ZStd3, 50)
    End If
    
    If Val(ZDesde4) <> 0 Or Val(ZZHasta4) <> 0 Then
        valor4.Text = Trim(ZDesde4) + " - " + Trim(ZHasta4) + " " + Trim(ZDescri4) + " " + Left$(ZStd4, 50)
            Else
        valor4.Text = Left$(ZStd4, 50)
    End If
    
    If Val(ZDesde5) <> 0 Or Val(ZZHasta5) <> 0 Then
        valor5.Text = Trim(ZDesde5) + " - " + Trim(ZHasta5) + " " + Trim(ZDescri5) + " " + Left$(ZStd5, 50)
            Else
        valor5.Text = Left$(ZStd5, 50)
    End If
    
    If Val(ZDesde6) <> 0 Or Val(ZZHasta6) <> 0 Then
        valor6.Text = Trim(ZDesde6) + " - " + Trim(ZHasta6) + " " + Trim(ZDescri6) + " " + Left$(ZStd6, 50)
            Else
        valor6.Text = Left$(ZStd6, 50)
    End If
    
    If Val(ZDesde7) <> 0 Or Val(ZZHasta7) <> 0 Then
        valor7.Text = Trim(ZDesde7) + " - " + Trim(ZHasta7) + " " + Trim(ZDescri7) + " " + Left$(ZStd7, 50)
            Else
        valor7.Text = Left$(ZStd7, 50)
    End If
    
    If Val(ZDesde8) <> 0 Or Val(ZZHasta8) <> 0 Then
        valor8.Text = Trim(ZDesde8) + " - " + Trim(ZHasta8) + " " + Trim(ZDescri8) + " " + Left$(ZStd8, 50)
            Else
        valor8.Text = Left$(ZStd8, 50)
    End If
    
    If Val(ZDesde9) <> 0 Or Val(ZZHasta9) <> 0 Then
        valor9.Text = Trim(ZDesde9) + " - " + Trim(ZHasta9) + " " + Trim(ZDescri9) + " " + Left$(ZStd9, 50)
            Else
        valor9.Text = Left$(ZStd9, 50)
    End If
    
    If Val(ZDesde10) <> 0 Or Val(ZZHasta10) <> 0 Then
        valor10.Text = Trim(ZDesde10) + " - " + Trim(ZHasta10) + " " + Trim(ZDescri10) + " " + Left$(ZStd10, 50)
            Else
        valor10.Text = Left$(ZStd10, 50)
    End If
    
    
    
    
    Call Conecta_Empresa
        
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo WError
    
    If WGraba <> "S" Then
    
        Call Ingresa_clave

               Else

        If Producto.Text <> "" Then
        
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
        
            Producto.Text = UCase(Producto.Text)
            Cliente.Text = UCase(Cliente.Text)
        
            WProducto = Producto.Text
            WCliente = Cliente.Text
            WClave = WProducto + WCliente

            WOpcion1 = Str$(Opcion1.Value)
            WOpcion2 = Str$(Opcion2.Value)
            WOpcion3 = Str$(Opcion3.Value)
            WOpcion4 = Str$(Opcion4.Value)
            WOpcion5 = Str$(Opcion5.Value)
            WOpcion6 = Str$(Opcion6.Value)
            WOpcion7 = Str$(Opcion7.Value)
            WOpcion8 = Str$(Opcion8.Value)
            WOpcion9 = Str$(Opcion9.Value)
            WOpcion10 = Str$(Opcion10.Value)
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM AltaCertificado"
            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
            ZSql = ZSql & " and AltaCertificado.Cliente = " + "'" + Cliente.Text + "'"
            spAltaCertificado = ZSql
            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
            If rstAltaCertificado.RecordCount > 0 Then
            
                rstAltaCertificado.Close
            
                ZSql = ""
                ZSql = ZSql & "UPDATE AltaCertificado SET "
                ZSql = ZSql & "Opcion1 = " + "'" + WOpcion1 + "',"
                ZSql = ZSql & "Opcion2 = " + "'" + WOpcion2 + "',"
                ZSql = ZSql & "Opcion3 = " + "'" + WOpcion3 + "',"
                ZSql = ZSql & "Opcion4 = " + "'" + WOpcion4 + "',"
                ZSql = ZSql & "Opcion5 = " + "'" + WOpcion5 + "',"
                ZSql = ZSql & "Opcion6 = " + "'" + WOpcion6 + "',"
                ZSql = ZSql & "Opcion7 = " + "'" + WOpcion7 + "',"
                ZSql = ZSql & "Opcion8 = " + "'" + WOpcion8 + "',"
                ZSql = ZSql & "Opcion9 = " + "'" + WOpcion9 + "',"
                ZSql = ZSql & "Opcion10 = " + "'" + WOpcion10 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                ZSql = ZSql & " and Cliente = " + "'" + WCliente + "'"
                        
                spAltaCertificado = ZSql
                Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
                    
                ZSql = ""
                ZSql = ZSql & "INSERT INTO AltaCertificado ("
                ZSql = ZSql & "Clave, "
                ZSql = ZSql & "Producto , "
                ZSql = ZSql & "Cliente , "
                ZSql = ZSql & "Opcion1 , "
                ZSql = ZSql & "Opcion2 , "
                ZSql = ZSql & "Opcion3 , "
                ZSql = ZSql & "Opcion4 , "
                ZSql = ZSql & "Opcion5 , "
                ZSql = ZSql & "Opcion6 , "
                ZSql = ZSql & "Opcion7 , "
                ZSql = ZSql & "Opcion8 , "
                ZSql = ZSql & "Opcion9 , "
                ZSql = ZSql & "Opcion10) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + WClave + "',"
                ZSql = ZSql & "'" + WProducto + "',"
                ZSql = ZSql & "'" + WCliente + "',"
                ZSql = ZSql & "'" + WOpcion1 + "',"
                ZSql = ZSql & "'" + WOpcion2 + "',"
                ZSql = ZSql & "'" + WOpcion3 + "',"
                ZSql = ZSql & "'" + WOpcion4 + "',"
                ZSql = ZSql & "'" + WOpcion5 + "',"
                ZSql = ZSql & "'" + WOpcion6 + "',"
                ZSql = ZSql & "'" + WOpcion7 + "',"
                ZSql = ZSql & "'" + WOpcion8 + "',"
                ZSql = ZSql & "'" + WOpcion9 + "',"
                ZSql = ZSql & "'" + WOpcion10 + "')"
           
                spAltaCertificado = ZSql
                Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            Call Conecta_Empresa
            
            Call CmdLimpiar_Click
            Producto.SetFocus
        
        End If
        
    End If
    
    Exit Sub

WError:
     Resume Next
    
End Sub

Private Sub cmdDelete_Click()

    If Producto.Text <> "" Then
    
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
    
        ZSql = ""
        ZSql = ZSql & "Select *"
        ZSql = ZSql & " FROM AltaCertificado"
        ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
        ZSql = ZSql & " and AltaCertificado.Cliente = " + "'" + Cliente.Text + "'"
        spAltaCertificado = ZSql
        Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
        If rstAltaCertificado.RecordCount > 0 Then
            rstAltaCertificado.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql & "DELETE AltaCertificado"
                ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
                ZSql = ZSql & " and AltaCertificado.Cliente = " + "'" + Cliente.Text + "'"
                spAltaCertificado = ZSql
                Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
        Call Conecta_Empresa
        
    End If
    
    Producto.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -     -   "
    Cliente.Text = ""
    DesProducto.Caption = ""
    DesCliente.Caption = ""
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    Opcion10.Value = 0
    Ensayo1.Text = ""
    Ensayo2.Text = ""
    Ensayo3.Text = ""
    Ensayo4.Text = ""
    Ensayo5.Text = ""
    Ensayo6.Text = ""
    Ensayo7.Text = ""
    Ensayo8.Text = ""
    Ensayo9.Text = ""
    Ensayo10.Text = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Valor1.Text = ""
    valor2.Text = ""
    Valor3.Text = ""
    valor4.Text = ""
    valor5.Text = ""
    valor6.Text = ""
    valor7.Text = ""
    valor8.Text = ""
    valor9.Text = ""
    valor10.Text = ""
    Valor11.Text = ""
    Valor22.Text = ""
    Valor33.Text = ""
    Valor44.Text = ""
    Valor55.Text = ""
    Valor66.Text = ""
    Valor77.Text = ""
    Valor88.Text = ""
    Valor99.Text = ""
    Valor1010.Text = ""
    WGraba = ""
    WGrabaII = ""
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgAltaCertificado.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    XIndice = 0

    Dim IngresaItem As String
    pantalla.Clear
    WIndice.Clear
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    spClientes = "ListaClienteConsulta"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then

    With rstClientes
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstClientes!Cliente + " " + rstClientes!razon
                pantalla.AddItem IngresaItem
                IngresaItem = rstClientes!Cliente
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstClientes.Close

    End If
    
    pantalla.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    WEspacios = Len(Ayuda.Text)
    WIndice.Clear
    
    XIndice = 0
    
    Select Case XIndice
        Case 0
            pantalla.Clear
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            Da = Len(rstCliente!razon) - WEspacios
                
                            For aa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!razon, aa, WEspacios) Then
                                    Auxi = rstCliente!Cliente
                                    IngresaItem = Auxi + "    " + rstCliente!razon
                                    pantalla.AddItem IngresaItem
                                    IngresaItem = rstCliente!Cliente
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
                rstCliente.Close
            End If
            
        
        Case Else
        
    End Select
    
    End If

End Sub


Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
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
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgAltaCertificado.Caption = "Ingreso de Datos a Imprimir en las Certificados de Analisis:  " + !Nombre
        End If
    End With
        
    Producto.Text = "  -     -   "
    Cliente.Text = ""
    DesProducto.Caption = ""
    DesCliente.Caption = ""
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    Opcion6.Value = 0
    Opcion7.Value = 0
    Opcion8.Value = 0
    Opcion9.Value = 0
    Opcion10.Value = 0
        
    WGraba = ""
    WGrabaII = ""
    EmpresaActual = WEmpresa
    
    If ZZPasaCliente <> "" And ZZPasaTerminado <> "" Then
        Producto.Text = ZZPasaTerminado
        Cliente.Text = ZZPasaCliente
        Call Producto_Keypress(13)
        Call Cliente_Keypress(13)
    End If
    
End Sub

Private Sub pantalla_Click()
    Indice = pantalla.ListIndex
    Cliente.Text = WIndice.List(Indice)
    Call Cliente_Keypress(13)
    pantalla.Visible = False
    Ayuda.Visible = False
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Terminado"
            ZSql = ZSql & " Where Terminado.Codigo = " + "'" + Producto.Text + "'"
            spTerminado = ZSql
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesProducto.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                    Else
                DesProducto.Caption = ""
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
  Rem BY NAN 7-7-2015
            ZSql = ""
            ZSql = ZSql & "Select producto "
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Producto.Text + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
                rstEspecifUnifica.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                Cliente.SetFocus
                    Else
                Call Conecta_Empresa
                WProducto = Producto.Text
                WCliente = Cliente.Text
                CmdLimpiar_Click
                Producto.Text = WProducto
                Cliente.Text = WCliente
                Producto.SetFocus
            End If
            
        End If
    End If
End Sub


Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM Cliente"
            ZSql = ZSql & " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!razon
                rstCliente.Close
                    Else
                DesCliente.Caption = ""
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
            
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM AltaCertificado"
            ZSql = ZSql & " Where AltaCertificado.Producto = " + "'" + Producto.Text + "'"
            ZSql = ZSql & " and AltaCertificado.cliente = " + "'" + Cliente.Text + "'"
            spAltaCertificado = ZSql
            Set rstAltaCertificado = db.OpenRecordset(spAltaCertificado, dbOpenSnapshot, dbSQLPassThrough)
        
            If rstAltaCertificado.RecordCount > 0 Then
                Opcion1.Value = rstAltaCertificado!Opcion1
                Opcion2.Value = rstAltaCertificado!Opcion2
                Opcion3.Value = rstAltaCertificado!Opcion3
                Opcion4.Value = rstAltaCertificado!Opcion4
                Opcion5.Value = rstAltaCertificado!Opcion5
                Opcion6.Value = rstAltaCertificado!Opcion6
                Opcion7.Value = rstAltaCertificado!Opcion7
                Opcion8.Value = rstAltaCertificado!Opcion8
                Opcion9.Value = rstAltaCertificado!Opcion9
                Opcion10.Value = rstAltaCertificado!Opcion10
                rstAltaCertificado.Close
            End If
            
            Call Conecta_Empresa
            
        End If
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
        If UCase(WClave.Text) = "PROCESO01" Then
            WGraba = "S"
            XClave.Visible = False
            Call cmdAdd_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Composicion de Productos")
            WClave.SetFocus
        End If
    End If
End Sub

