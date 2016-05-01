VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgIngresosVarios 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Valores Varios"
   ClientHeight    =   8010
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   11880
   Begin VB.CommandButton Impresion 
      Caption         =   "Impres. F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8520
      MouseIcon       =   "ingresosvarios.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Impresion de Orden de Pago"
      Top             =   6480
      Width           =   855
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
      Left            =   7680
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
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
      Height          =   1035
      Left            =   8640
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
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
      Height          =   1620
      ItemData        =   "ingresosvarios.frx":0B4C
      Left            =   7680
      List            =   "ingresosvarios.frx":0B53
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Cheque 
      Caption         =   "Cheques "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MouseIcon       =   "ingresosvarios.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":0E6B
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Cartera de Cheques"
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Tipo1 
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
      Left            =   8040
      MaxLength       =   8
      TabIndex        =   38
      Text            =   " "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Numero1 
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
      Left            =   8640
      MaxLength       =   8
      TabIndex        =   37
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      MouseIcon       =   "ingresosvarios.frx":12FA
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":1604
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Primer Registro"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MouseIcon       =   "ingresosvarios.frx":1A46
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":1D50
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Registro Anterior"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguien. F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      MouseIcon       =   "ingresosvarios.frx":2192
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":249C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Registro Siguiente"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      MouseIcon       =   "ingresosvarios.frx":28DE
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":2BE8
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Salida"
      Top             =   6480
      Width           =   855
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Cuenta 
      BeginProperty Font 
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
      MaxLength       =   10
      TabIndex        =   29
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MouseIcon       =   "ingresosvarios.frx":302A
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":3334
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Menu Principal"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MouseIcon       =   "ingresosvarios.frx":3B76
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":3E80
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Consulta de Datos"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpia F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      MouseIcon       =   "ingresosvarios.frx":46C2
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":49CC
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10560
      MouseIcon       =   "ingresosvarios.frx":520E
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":5518
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Elimina el Registro"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      MouseIcon       =   "ingresosvarios.frx":5D5A
      MousePointer    =   99  'Custom
      Picture         =   "ingresosvarios.frx":6064
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2160
      Width           =   855
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4080
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4200
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3600
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
      TabIndex        =   15
      Top             =   3600
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
      TabIndex        =   14
      Top             =   3600
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
      Left            =   2760
      TabIndex        =   13
      Top             =   3120
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Top             =   3600
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
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin Crystal.CrystalReport listado 
      Left            =   5880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
   End
   Begin VB.TextBox Observaciones 
      BeginProperty Font 
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
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   4695
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox Recibo 
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
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Top             =   3120
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
      Height          =   4335
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label2 
      Caption         =   "Comprobante"
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
      Left            =   6600
      TabIndex        =   36
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Cuenta Contable"
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
      TabIndex        =   30
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DesCuenta 
      BackColor       =   &H00C0C000&
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
      Left            =   3360
      TabIndex        =   28
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "VALORES RECIBIDOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
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
      Left            =   5040
      TabIndex        =   9
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  1) Ef.   2) Ch.   3) Doc. 5) Recha."
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
      Left            =   1680
      TabIndex        =   8
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero"
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
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "PrgIngresosVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double
Private ZSaldo As Double
Dim ZMes As String
Dim ZAno As String
Dim ZRecibo As String
Dim Numero As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Suma_Datos()

    Creditos.Caption = ""
    For iRow = 1 To 100
        WCreditos = Val(WVector1.TextMatrix(iRow, 5))
        Creditos.Caption = Str$(Val(Creditos.Caption) + WCreditos)
    Next iRow
    
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    
End Sub

Private Sub Lee_Datos()

    Tipo1.Text = ""
    Numero1.Text = ""

    Call Limpia_Vector

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Do
        With rstRecibos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", ZRecibo + Auxi1
            If .NoMatch = False Then
                Select Case Val(!Tiporeg)
                    Case 1
                        If Val(!Tipo1) <> 0 And Val(!Tipo1) <> 99 Then
                            Tipo1.Text = !Tipo1
                            Numero1.Text = !Numero1
                        End If
                    Case 2
                        Credito = Credito + 1
                        WVector1.Row = Credito
                        WVector1.Col = 1
                        WVector1.Text = !Tipo2
                        WVector1.Col = 2
                        WVector1.Text = !Numero2
                        WVector1.Col = 3
                        WVector1.Text = !Fecha2
                        WVector1.Col = 4
                        WVector1.Text = !Banco2
                        WVector1.Col = 5
                        WVector1.Text = !Importe2
                        WVector1.Text = Alinea("###,###.##", WVector1.Text)
                        WVector1.Col = 6
                        WVector1.Text = !Estado2
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()
End Sub

Private Sub Anterior_Click()
    Auxi1 = Str$(Val(Recibo.Text) + 500000)
    Call Ceros(Auxi1, 6)
    With rstRecibos
        .Index = "Clave"
        .Seek "<", Auxi1 + "00"
        If .NoMatch = False Then
            If !Recibo > 500000 Then
                Recibo.Text = Mid$(Str$((!Recibo - 500000)), 2, 6)
                    Else
                Recibo.Text = ""
            End If
                Else
            Recibo.Text = ""
        End If
    End With
    Call Recibo_Keypress(13)
End Sub

Private Sub Cheque_Click()

    Opcion.Clear
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cheques Rechazados"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub cmdAdd_Click()

    If Val(Recibo.Text) = 0 Then
        With rstRecibos
            .Index = "Clave"
            .Seek "<", "99999999"
            If .NoMatch = False Then
                If !Recibo > 500000 Then
                    Recibo.Text = Mid$(Str$((!Recibo - 500000) + 1), 2, 6)
                        Else
                    Recibo.Text = "1"
                End If
                    Else
                Recibo.Text = "1"
            End If
        End With
    End If

    If Recibo.Text <> "" And Fecha.Text <> "" Then
    
        ZRecibo = Str$(Val(Recibo.Text) + 500000)
        Call Ceros(ZRecibo, 6)
    
        Auxi1 = ZRecibo
        Call Ceros(Auxi1, 6)
        ZRecibo = Auxi1
        
        With rstRecibos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = ZRecibo + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
            
                ZRecibo = Str$(Val(Recibo.Text) + 500000)
                Call Ceros(ZRecibo, 6)
                
                With rstRecibos
                    For da = 1 To 99
                        Auxi1 = Str$(da)
                        Call Ceros(Auxi1, 2)
                        .Index = "Clave"
                        .Seek "=", ZRecibo + Auxi1
                        If .NoMatch = False Then
                            .Delete
                        End If
                    Next da
                End With
                
            End If
        End With
    
        If Existe <> "S" Then
    
            With rstRecibos
                Renglon = 0
                .Index = "Clave"
                For iRow = 1 To 100
        
                    WVector1.Col = 5
                    WVector1.Row = iRow
                    If Val(WVector1.Text) <> 0 Then
                        .AddNew
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        !Recibo = ZRecibo
                        !Renglon = Auxi1
                        !Cliente = 0
                        !Fecha = Fecha.Text
                        !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        !TipoRec = "3"
                        !Retganancias = 0
                        !RetIva = 0
                        !RetOtra = 0
                        !RetSuss = 0
                        !NroRetganancias = 0
                        !NroRetIva = 0
                        !NroRetOtra = 0
                        !NroRetSuss = 0
                        !Retencion = 0
                        !Tiporeg = "2"
                        !Tipo1 = ""
                        !Letra1 = ""
                        !Punto1 = ""
                        !Numero1 = ""
                        !Importe1 = 0
                        WVector1.Col = 1
                        !Tipo2 = WVector1.Text
                        WVector1.Col = 2
                        !Numero2 = WVector1.Text
                        WVector1.Col = 3
                        !Fecha2 = WVector1.Text
                        !FechaOrd2 = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                        WVector1.Col = 4
                        !Banco2 = WVector1.Text
                        WVector1.Col = 5
                        !Importe2 = Val(WVector1.Text)
                        WVector1.Col = 6
                        !Estado2 = Trim(WVector1.Text)
                        If !Estado2 = "" Then
                            !Estado2 = "P"
                        End If
                        !Observaciones = Observaciones.Text
                        !Empresa = Val(WEmpresa)
                        !Clave = !Recibo + !Renglon
                        !Importe = Credito
                        If !Tipo2 = 4 Then
                            !Cuenta = WCuenta(iRow)
                                Else
                            !Cuenta = "1"
                        End If
                        !Destino = ""
                        !Orden = 0
                        !Deposito = 0
                        .Update
                        .Bookmark = .LastModified
                    
                    End If
            
                Next iRow
            End With
        
            With rstRecibos
                .AddNew
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                !Recibo = ZRecibo
                !Renglon = Auxi1
                !Cliente = 0
                !Fecha = Fecha.Text
                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !TipoRec = "3"
                !Retganancias = 0
                !RetIva = 0
                !RetOtra = 0
                !RetSuss = 0
                !NroRetganancias = 0
                !NroRetIva = 0
                !NroRetOtra = 0
                !NroRetSuss = 0
                !Retencion = 0
                !Tiporeg = "1"
                If Val(Tipo1.Text) <> 0 Then
                    !Tipo1 = Tipo1.Text
                        Else
                    !Tipo1 = "99"
                End If
                !Letra1 = ""
                !Punto1 = ""
                If Val(Numero1.Text) <> 0 Then
                    !Numero1 = Numero1.Text
                        Else
                    !Numero1 = ZRecibo
                End If
                !Importe1 = Credito
                !Tipo2 = ""
                !Numero2 = ""
                !Fecha2 = ""
                !FechaOrd2 = ""
                !Banco2 = ""
                !Importe2 = 0
                !Estado2 = ""
                !Observaciones = Observaciones.Text
                !Empresa = Val(WEmpresa)
                !Clave = !Recibo + !Renglon
                !Importe = Credito
                !Cuenta = ""
                !Destino = ""
                !Orden = 0
                !Deposito = 0
                .Update
                .Bookmark = .LastModified
            End With
            
            If Val(Tipo1.Text) <> 0 And Val(Numero1.Text) <> 0 Then
                With rstRecibos
                    .Index = "Fecha2"
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(!Tiporeg) = 2 Then
                                If !Tipo2 = Tipo1.Text And !Numero2 = Numero1.Text Then
                                    .Edit
                                    !Estado2 = "X"
                                    .Update
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
            End If
            
            mm$ = "El numero de movimiento es " + Recibo.Text
            A% = MsgBox(mm$, 0, "Ingreso de Dinero")
            
            Call ImpreComprobante
                
            Call CmdLimpiar_Click
            Recibo.SetFocus
                
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Recibo.Text <> "" Then
    
        T$ = "Ingresos de Recibos"
        m1$ = "Desea Anular el Recibo"
        Respuesta% = MsgBox(m1$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            ZRecibo = Str$(Val(Recibo.Text) + 500000)
            Call Ceros(ZRecibo, 6)
            
            With rstRecibos
                For da = 1 To 99
                    Auxi1 = Str$(da)
                    Call Ceros(Auxi1, 2)
                    .Index = "Clave"
                    .Seek "=", ZRecibo + Auxi1
                    If .NoMatch = False Then
                        .Delete
                    End If
                Next da
            End With
                    
            Call CmdLimpiar_Click
                        
        End If
        
    End If
    
    Recibo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector

    Recibo.Text = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Recibo.SetFocus
    Creditos.Caption = ""
    Tipo1.Text = ""
    Numero1.Text = ""
    
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
    
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstRecibos
        .Close
    End With
    DbsAdminis.Close
    PrgIngresosVarios.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub


Private Sub Impresion_Click()
    Call ImpreComprobante
End Sub

Private Sub ImpreComprobante()

    T$ = "Impresion de Comprobante del Ingreso"
    m1$ = "Desea realizar la impresion del comprobante"
    Respuesta% = MsgBox(m1$, 32 + 4, T$)
    If Respuesta% = 6 Then

        listado.GroupSelectionFormula = "{Recibos.Recibo} in " + Chr$(34) + ZRecibo + Chr$(34) + " to " + Chr$(34) + ZRecibo + Chr$(34)
        listado.SelectionFormula = "{Recibos.Recibo} in " + Chr$(34) + ZRecibo + Chr$(34) + " to " + Chr$(34) + ZRecibo + Chr$(34)
        
        listado.ReportFileName = "ImpreIngVar.rpt"
        listado.DataFiles(0) = WEmpresa + "admi.mdb"
    
        listado.Destination = 1
        Rem listado.Destination = 0
        listado.Action = 1

    End If

End Sub


Private Sub Primer_Click()
    With rstRecibos
        .Index = "Clave"
        .Seek ">", "50000000"
        If .NoMatch = False Then
            If !Recibo > 500000 Then
                Recibo.Text = Mid$(Str$((!Recibo - 500000)), 2, 6)
                    Else
                Recibo.Text = ""
            End If
                Else
            Recibo.Text = ""
        End If
    End With
    Call Recibo_Keypress(13)
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZRecibo = Str$(Val(Recibo.Text) + 500000)
        Call Ceros(ZRecibo, 6)
    
        Auxi1 = ZRecibo
        Call Ceros(Auxi1, 6)
        ZRecibo = Auxi1
        
        With rstRecibos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = ZRecibo + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                Observaciones.Text = !Observaciones
                Fecha.Text = !Fecha
            End If
        End With
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Recibo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem WVector1.Col = 1
        Rem WVector1.Row = 1
        Rem Call StartEdit
        Cuenta.SetFocus
    End If
    If KeyAscii = 27 Then
       Observaciones.Text = ""
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
            With rstCuenta
                .Index = "Cuenta"
                .Seek "=", Cuenta.Text
                If .NoMatch = False Then
                    WVector1.Col = 1
                    WVector1.Row = 1
                    Call StartEdit
                        Else
                    Cuenta.SetFocus
                End If
            End With
                Else
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear

    Opcion.AddItem "Cuentas Contables"

    Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstCuenta
                .Index = "Descripcion"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cuenta + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cuenta
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Ayuda.SetFocus
            
        Case 1
            With rstRecibos
                .Index = "Fecha2"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tiporeg) = 2 Then
                            If (Val(!Tipo2) = 3 Or Val(!Tipo2) = 5) And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                Numero = Str$(Val(!Numero2))
                                Call Ceros(Numero, 6)
                                IngresaItem = Numero + "    " + !Fecha2 + "      " + Auxi$ + "      " + !Banco2
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Tipo2 + !Numero2
                                WIndice.AddItem IngresaItem
                                IngresaItem = ""
                                WVector.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 2
            With rstBanco
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If !Banco >= 100 Then
                            Auxi1 = Str$(!Banco)
                            Call Ceros(Auxi1, 4)
                            IngresaItem = Auxi1 + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Banco
                            WIndice.AddItem IngresaItem
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Select Case XIndice
        Case 0
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Cuenta.Text = WIndice.List(Indice)
            Cuenta.SetFocus
            
        Case 1
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Tipo1.Text = Left$(WIndice.List(Indice), 2)
            Numero1.Text = Right$(WIndice.List(Indice), 8)
            
        Case 2
            WTexto1.Visible = False
            WTexto2.Visible = False
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            WBanco = WIndice.List(Indice)
            With rstBanco
                .Index = "Banco"
                .Seek "=", WBanco
                If .NoMatch = False Then
                    WVector1.Text = !Nombre
                End If
            End With
            Call StartEdit
            
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstCuenta
                .Index = "Descripcion"
                .MoveFirst
                Do
                    If .EOF = False Then
            
                        da = Len(!Descripcion) - WEspacios
                
                        For aa = 1 To da + 1
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                IngresaItem = !Cuenta + "    " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Cuenta
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
            
        Case 1
            With rstBanco
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If !Banco >= 100 Then
                            da = Len(!Nombre) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                    IngresaItem = Str$(!Banco) + " " + !Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Banco
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case Else
    End Select
    
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Activate()
    PrgIngresosVarios.Caption = "Ingreso de Valores Varios : " + WNombreEmpresa
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Recibos
    OPEN_FILE_Cuenta
    OPEN_FILE_Banco
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Recibo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    Tipo1.Text = ""
    Numero1.Text = ""
    
End Sub


Rem
Rem Controles de la grilla
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

Private Sub TotalRete_DblClick()
    PantaRete.Visible = True
    Retganancias.SetFocus
End Sub

Private Sub Siguiente_Click()
    Auxi1 = Str$(Val(Recibo.Text) + 500000)
    Call Ceros(Auxi1, 6)
    With rstRecibos
        .Index = "Clave"
        .Seek ">", Auxi1 + "99"
        If .NoMatch = False Then
            If !Recibo > 500000 Then
                Recibo.Text = Mid$(Str$((!Recibo - 500000)), 2, 6)
                    Else
                Recibo.Text = ""
            End If
                Else
            Recibo.Text = ""
        End If
    End With
    Call Recibo_Keypress(13)
End Sub

Private Sub Ultimo_Click()
    With rstRecibos
        .Index = "Clave"
        .Seek "<", "99999999"
        If .NoMatch = False Then
            If !Recibo > 500000 Then
                Recibo.Text = Mid$(Str$((!Recibo - 500000)), 2, 6)
                    Else
                Recibo.Text = ""
            End If
                Else
            Recibo.Text = ""
        End If
    End With
    Call Recibo_Keypress(13)
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
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

Private Sub Control_Grilla()
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


Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
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

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
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
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Numero/Cta"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Banco/Cliente"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case Else
                
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
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub
    
Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 5 Then
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 2)
                WVector1.Text = Auxi$
                Select Case Val(WVector1.Text)
                    Case 1
                        WVector1.Col = 2
                        WVector1.Text = ""
                        WVector1.Col = 3
                        WVector1.Text = ""
                        WVector1.Col = 4
                        WVector1.Text = ""
                    Case Else
                End Select
                    Else
                WControl = "N"
            End If
            
        Case 2
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
        Case 3
            Call Valida_fecha1(WVector1.Text, Auxi)
            If Auxi <> "S" Then
                WControl = "N"
            End If
            
        Case 5
            iRow = WVector1.Row
            WVector1.Col = 1
            XTipo = WVector1.Text
            WVector1.Col = 5
            WVector1.Text = Alinea("###,###.##", WVector1.Text)
            Call Suma_Datos
            WVector1.Row = iRow
        
        Case Else
    End Select
End Sub


Private Sub Cuenta_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cuentas Contables"
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub WTexto1_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Contables"
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub WTexto2_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Contables"
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Recibo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call cmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub



Private Sub WVector1_DblClick()

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
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
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
    
End Sub















