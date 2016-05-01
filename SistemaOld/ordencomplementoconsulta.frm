VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdenComplementoConsulta 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de datos Complementarios de la Orden de Compra"
   ClientHeight    =   6585
   ClientLeft      =   2055
   ClientTop       =   2565
   ClientWidth     =   11685
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.Frame PantaAlta 
      BackColor       =   &H00C0C0C0&
      Height          =   3375
      Left            =   -720
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   8895
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
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   39
         Text            =   " "
         Top             =   600
         Width           =   6855
      End
      Begin VB.TextBox ObservacionesI 
         BeginProperty Font 
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
         MaxLength       =   80
         TabIndex        =   38
         Text            =   " "
         Top             =   960
         Width           =   6855
      End
      Begin VB.TextBox ObservacionesII 
         BeginProperty Font 
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
         MaxLength       =   80
         TabIndex        =   37
         Text            =   " "
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox ObservacionesIII 
         BeginProperty Font 
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
         MaxLength       =   80
         TabIndex        =   36
         Text            =   " "
         Top             =   1680
         Width           =   6855
      End
      Begin VB.TextBox ObservacionesIV 
         BeginProperty Font 
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
         MaxLength       =   80
         TabIndex        =   35
         Text            =   " "
         Top             =   2040
         Width           =   6855
      End
      Begin VB.CommandButton CierraPanta 
         Caption         =   "Graba"
         Height          =   615
         Left            =   2640
         TabIndex        =   34
         Top             =   2520
         Width           =   2895
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1800
         TabIndex        =   40
         Top             =   240
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
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
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
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame IngresaObservaciones 
      Caption         =   "Ingreso de Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.CommandButton AltaObserva 
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
         Height          =   495
         Left            =   6600
         TabIndex        =   32
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton FinIngresaObserva 
         Caption         =   "Fin de Ingreso"
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
         Left            =   4920
         TabIndex        =   4
         Top             =   5160
         Width           =   1575
      End
      Begin VB.ComboBox EntregaII 
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
         Left            =   2640
         TabIndex        =   29
         Top             =   5400
         Width           =   2175
      End
      Begin VB.ComboBox EntregaI 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   5400
         Width           =   2295
      End
      Begin VB.ComboBox PagoLetra 
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   5400
         Width           =   2655
      End
      Begin VB.TextBox ImpoLetra 
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
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   19
         Text            =   " "
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox ImpoDespacho 
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
         Left            =   8760
         MaxLength       =   20
         TabIndex        =   18
         Text            =   " "
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Orden 
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
         MaxLength       =   6
         TabIndex        =   14
         Text            =   " "
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Carpeta 
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
         MaxLength       =   6
         TabIndex        =   13
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox PagoDespacho 
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3000
         Width           =   2655
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
         TabIndex        =   8
         Top             =   1080
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   3960
         TabIndex        =   6
         Top             =   720
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
         Left            =   2520
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   390
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
         Left            =   3360
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   4575
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8070
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox FechaLlegada 
         Height          =   285
         Left            =   8760
         TabIndex        =   9
         Top             =   1440
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
      Begin MSMask.MaskEdBox VtoDespacho 
         Height          =   285
         Left            =   5880
         TabIndex        =   15
         Top             =   5460
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox VtoLetra 
         Height          =   285
         Left            =   8760
         TabIndex        =   21
         Top             =   4620
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
      Begin MSMask.MaskEdBox VtoLetraII 
         Height          =   285
         Left            =   8760
         TabIndex        =   25
         Top             =   4620
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox FechaEmbarque 
         Height          =   285
         Left            =   8760
         TabIndex        =   43
         Top             =   720
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha de Embarque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   44
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Documento de Embarque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   30
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Certificado de Analisis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   5040
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   " Vencimiento de la Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8280
         TabIndex        =   26
         Top             =   4200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Estado de la Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   24
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fecha Prevista de Pago de la Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8280
         TabIndex        =   23
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Importe de la Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   22
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Importe del Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   17
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Fecha Prevista de Pago del Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   16
         Top             =   5040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Estado del Despacho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha Prevista de Arribo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
      End
   End
End
Attribute VB_Name = "PrgOrdenComplementoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstObservaOrden As Recordset
Dim spObservaOrden As String
Dim EmpresaAnterior As String
Dim EmpresaOrden As Integer

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String

Dim WControl As String

Private Sub cmdClose_Click()
    PrgOrdenComplemento.Hide
    Unload Me
    PrgOrden.Show
End Sub

Rem
Rem Controles de la WVector1
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

Private Sub AltaObserva_Click()
    Fecha.Text = "  /  /    "
    Observaciones.Text = ""
    ObservacionesI.Text = ""
    ObservacionesII.Text = ""
    ObservacionesIII.Text = ""
    ObservacionesIV.Text = ""
    PantaAlta.Visible = True
    Fecha.SetFocus
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesI.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub ObservacionesI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesI.SetFocus
    End If
End Sub

Private Sub ObservacionesII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesII.SetFocus
    End If
End Sub

Private Sub ObservacionesIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObservacionesIV.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesIII.SetFocus
    End If
End Sub

Private Sub ObservacionesIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        ObservacionesIV.SetFocus
    End If
End Sub


Private Sub CierraPanta_Click()

    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    WRenglon = 0
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrden"
    Sql3 = " Where ObservaOrden.Carpeta = " + "'" + Carpeta.Text + "'"
    Sql4 = " Order by ObservaOrden.Clave"
    spObservaOrden = Sql1 + Sql2 + Sql3 + Sql4
    Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrden.RecordCount > 0 Then
        With rstObservaOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstObservaOrden.Close
    End If
    
    
    Dim ZZEntraObs(100, 10) As String
    
    Erase ZZEntraObs
    
    ZZEntraObs(1, 1) = Fecha.Text
    ZZEntraObs(1, 2) = Observaciones.Text
    
    ZZEntraObs(2, 1) = ""
    ZZEntraObs(2, 2) = ObservacionesI.Text
    
    ZZEntraObs(3, 1) = ""
    ZZEntraObs(3, 2) = ObservacionesII.Text
    
    ZZEntraObs(4, 1) = ""
    ZZEntraObs(4, 2) = ObservacionesIII.Text
    
    ZZEntraObs(5, 1) = ""
    ZZEntraObs(5, 2) = ObservacionesIV.Text
    
    For ZZCiclo = 1 To 5
    
        If Trim(ZZEntraObs(ZZCiclo, 2)) <> "" Then
    
            WCarpeta = Carpeta.Text
            Auxi = Carpeta.Text
            Call Ceros(Auxi, 6)
            
            WOrden = Orden.Text
            
            WRenglon = WRenglon + 1
            Auxi1 = Str$(WRenglon)
            Call Ceros(Auxi1, 2)
            
            WClave = Auxi + Auxi1
            
            WTexto1 = ZZEntraObs(ZZCiclo, 1)
            WTexto2 = ZZEntraObs(ZZCiclo, 2)
            
            Sql1 = "INSERT INTO ObservaOrden ("
            Sql2 = "Clave ,"
            Sql3 = "Carpeta ,"
            Sql4 = "Renglon ,"
            Sql5 = "Orden ,"
            Sql6 = "Texto1 ,"
            Sql7 = "Texto2 )"
            Sql8 = "Values ("
            Sql9 = "'" + WClave + "',"
            Sql10 = "'" + WCarpeta + "',"
            Sql11 = "'" + Auxi1 + "',"
            Sql12 = "'" + WOrden + "',"
            Sql13 = "'" + WTexto1 + "',"
            Sql14 = "'" + WTexto2 + "')"
            
            spObservaOrden = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                     + Sql11 + Sql12 + Sql13 + Sql14
            Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
    Next ZZCiclo
    
    Call Limpia_VectorII
    WRenglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrden"
    Sql3 = " Where ObservaOrden.Carpeta = " + "'" + Carpeta.Text + "'"
    Sql4 = " Order by ObservaOrden.Clave"
    spObservaOrden = Sql1 + Sql2 + Sql3 + Sql4
    Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrden.RecordCount > 0 Then
        With rstObservaOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = IIf(IsNull(rstObservaOrden!texto1), "", rstObservaOrden!texto1)
            
                    WVector1.Col = 2
                    WVector1.Text = IIf(IsNull(rstObservaOrden!texto2), "", rstObservaOrden!texto2)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstObservaOrden.Close
    End If
    
    Call Conecta_Empresa
    
    PantaAlta.Visible = False

End Sub

Private Sub EntregaI_DBLCLICK()
    PrgOrdenComplementoCertificado.Show
End Sub

Private Sub FechaLlegada_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ImpoDespacho.SetFocus
    End If
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Orden = " + "'" + WPasaOrden + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
       Rem BY NAN
        
        EntregaI.ListIndex = IIf(IsNull(rstOrden!EntregaI), "0", rstOrden!EntregaI)
        rstOrden.Close
    End If

End Sub

Private Sub ImpoDespacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoDespacho.SetFocus
    End If
End Sub

Private Sub Label9_dblClick()
    PrgOrdenComplementoCertificado.Show
End Sub

Private Sub VtoDespacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PagoDespacho.SetFocus
    End If
End Sub

Private Sub PagoDespacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ImpoLetra.SetFocus
    End If
End Sub

Private Sub ImpoLetra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoLetra.SetFocus
    End If
End Sub

Private Sub VtoLetra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        VtoLetraII.SetFocus
    End If
End Sub

Private Sub VtoLetraII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PagoLetra.SetFocus
    End If
End Sub

Private Sub PagoLetra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaLlegada.SetFocus
    End If
End Sub

Private Sub Form_Load()

    Orden.Text = WPasaOrden
    Carpeta.Text = WPasaCarpeta

    PagoDespacho.Clear
    
    PagoDespacho.AddItem "Pendiente"
    PagoDespacho.AddItem "Efectuado"
    
    PagoDespacho.ListIndex = 0

    PagoLetra.Clear
    
    PagoLetra.AddItem "Pendiente"
    PagoLetra.AddItem "Efectuado"
    
    PagoLetra.ListIndex = 0

    EntregaI.Clear
    
    EntregaI.AddItem "Pendiente"
    EntregaI.AddItem "Entregado"
    
    EntregaI.ListIndex = 0

    EntregaII.Clear
    
    EntregaII.AddItem "Pendiente"
    EntregaII.AddItem "Entregado"
    
    EntregaII.ListIndex = 0
    
    Call Limpia_VectorII
    
    WRenglon = 0
    
    XEmpresa = Wempresa
    Select Case Val(Wempresa)
        Case 1, 3, 5, 6, 7, 10, 11
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        Case Else
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrden"
    Sql3 = " Where ObservaOrden.Carpeta = " + "'" + Carpeta.Text + "'"
    Sql4 = " Order by ObservaOrden.Clave"
    spObservaOrden = Sql1 + Sql2 + Sql3 + Sql4
    Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrden.RecordCount > 0 Then
        With rstObservaOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                
                    WVector1.Col = 1
                    WVector1.Text = IIf(IsNull(rstObservaOrden!texto1), "", rstObservaOrden!texto1)
            
                    WVector1.Col = 2
                    WVector1.Text = IIf(IsNull(rstObservaOrden!texto2), "", rstObservaOrden!texto2)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstObservaOrden.Close
    End If
    
    Call Conecta_Empresa
    
    Call Busca_Empresa
            
    XEmpresa = Wempresa
    Select Case EmpresaOrden
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
    
 Rem nan
 
    Auxi = Orden.Text
    Call Ceros(Auxi, 6)
    WClave = Auxi + "01"
            
    spOrden = "ConsultaOrden " + "'" + WClave + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        FechaEmbarque.Text = IIf(IsNull(rstOrden!FechaEmbarque), "  /  /    ", rstOrden!FechaEmbarque)
        FechaLlegada.Text = IIf(IsNull(rstOrden!FechaLlegada), "  /  /    ", rstOrden!FechaLlegada)
        PagoDespacho.ListIndex = IIf(IsNull(rstOrden!PagoDespacho), "0", rstOrden!PagoDespacho)
        Rem ImpoDespacho.Text = IIf(IsNull(rstOrden!ImpoDespacho), " ", Str$(rstOrden!ImpoDespacho))
        ImpoDespacho.Text = IIf(IsNull(rstOrden!ImpoDespacho), "0", Int(rstOrden!ImpoDespacho))
        VtoDespacho.Text = IIf(IsNull(rstOrden!VtoDespacho), "  /  /    ", rstOrden!VtoDespacho)
        PagoLetra.ListIndex = IIf(IsNull(rstOrden!PagoLetra), "0", rstOrden!PagoLetra)
        EntregaI.ListIndex = IIf(IsNull(rstOrden!EntregaI), "0", rstOrden!EntregaI)
        EntregaII.ListIndex = IIf(IsNull(rstOrden!EntregaII), "0", rstOrden!EntregaII)
        ImpoLetra.Text = IIf(IsNull(rstOrden!ImpoLetra), "0", Int(rstOrden!ImpoLetra))
        VtoLetra.Text = IIf(IsNull(rstOrden!VtoLetra), "  /  /    ", rstOrden!VtoLetra)
        VtoLetraII.Text = IIf(IsNull(rstOrden!VtoLetraII), "  /  /    ", rstOrden!VtoLetraII)
        rstOrden.Close
    End If
    
    Call Conecta_Empresa
    
    WVector1.Col = 1
    WVector1.Row = 1
    WVector1.TopRow = 1
    Rem Call StartEditII
    
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
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII

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
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII

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
            Call Control_CampoII
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEditII

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
    StartEditII
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

Private Sub StartEditII()
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
        Case 2
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
                WVector1.Col = 1
            End If
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            WVector1.Col = XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_VectorII()

    WVector1.Clear

    Rem ponga la WVector1 en negritas
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

    ' Establesco loa Valores de la WVector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
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
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 6000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 80
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector1
    
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

Private Sub Lee_Observaciones()
End Sub

Private Sub FinIngresaObserva_Click()

    XEmpresa = Wempresa
    Select Case EmpresaOrden
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

    Auxi = Orden.Text
    Call Ceros(Auxi, 6)
    WClave = Auxi + "01"
            
    spOrden = "ConsultaOrden " + "'" + WClave + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        ZZFechaEmbarque = IIf(IsNull(rstOrden!FechaEmbarque), "  /  /    ", rstOrden!FechaEmbarque)
        ZZFechaLlegada = IIf(IsNull(rstOrden!FechaLlegada), "  /  /    ", rstOrden!FechaLlegada)
        ZZPagoDespacho = IIf(IsNull(rstOrden!PagoDespacho), "0", rstOrden!PagoDespacho)
        ZZImpoDespacho = IIf(IsNull(rstOrden!ImpoDespacho), "0", Int(rstOrden!ImpoDespacho))
        ZZVtoDespacho = IIf(IsNull(rstOrden!VtoDespacho), "  /  /    ", rstOrden!VtoDespacho)
        ZZPagoLetra = IIf(IsNull(rstOrden!PagoLetra), "0", rstOrden!PagoLetra)
        ZZEntregaI = IIf(IsNull(rstOrden!EntregaI), "0", rstOrden!EntregaI)
        ZZEntregaII = IIf(IsNull(rstOrden!EntregaII), "0", rstOrden!EntregaII)
        ZZImpoLetra = IIf(IsNull(rstOrden!ImpoLetra), "0", Int(rstOrden!ImpoLetra))
        ZZVtoLetra = IIf(IsNull(rstOrden!VtoLetra), "  /  /    ", rstOrden!VtoLetra)
        ZZVtoLetraII = IIf(IsNull(rstOrden!VtoLetraII), "  /  /    ", rstOrden!VtoLetraII)
        rstOrden.Close
    End If

    ZFechaEmbarque = FechaEmbarque.Text
    ZFechaLlegada = FechaLlegada.Text
    ZPagoDespacho = PagoDespacho.ListIndex
    ZImpoDespacho = ImpoDespacho.Text
    ZVtoDespacho = VtoDespacho.Text
    ZPagoLetra = PagoLetra.ListIndex
    ZEntregaI = EntregaI.ListIndex
    ZEntregaII = EntregaII.ListIndex
    ZImpoLetra = ImpoLetra.Text
    ZVtoLetra = VtoLetra.Text
    ZVtoLetraII = VtoLetraII.Text
    
    ZZGraba = "N"
    
    If ZFechaEmbarque <> ZZFechaEmbarque Then
        ZZGraba = "S"
    End If
    If ZFechaLlegada <> ZZFechaLlegada Then
        ZZGraba = "S"
    End If
    If ZPagoDespacho <> ZZPagoDespacho Then
        ZZGraba = "S"
    End If
    If Val(ZImpoDespacho) <> ZZImpoDespacho Then
        ZZGraba = "S"
    End If
    If ZVtoDespacho <> ZZVtoDespacho Then
        ZZGraba = "S"
    End If
    If ZPagoLetra <> ZZPagoLetra Then
        ZZGraba = "S"
    End If
    If ZEntregaI <> ZZEntregaI Then
        ZZGraba = "S"
    End If
    If ZEntregaII <> ZZEntregaII Then
        ZZGraba = "S"
    End If
    If Val(ZImpoLetra) <> ZZImpoLetra Then
        ZZGraba = "S"
    End If
    If ZVtoLetra <> ZZVtoLetra Then
        ZZGraba = "S"
    End If
    If ZVtoLetraII <> ZZVtoLetraII Then
        ZZGraba = "S"
    End If
    
    If ZZGraba = "S" Then
        T$ = "Actualizacion de Orden de Compra"
        m$ = "Desea Guardar los cambios a la Orden de Compra "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Orden SET "
            ZSql = ZSql + " FechaEmbarque = " + "'" + ZFechaEmbarque + "',"
            ZSql = ZSql + " FechaLLegada = " + "'" + ZFechaLlegada + "',"
            ZSql = ZSql + " PagoDespacho = " + "'" + Str$(ZPagoDespacho) + "',"
            ZSql = ZSql + " ImpoDespacho = " + "'" + ZImpoDespacho + "',"
            ZSql = ZSql + " VtoDespacho = " + "'" + ZVtoDespacho + "',"
            ZSql = ZSql + " PagoLetra = " + "'" + Str$(ZPagoLetra) + "',"
            ZSql = ZSql + " EntregaI = " + "'" + Str$(ZEntregaI) + "',"
            ZSql = ZSql + " EntregaII = " + "'" + Str$(ZEntregaII) + "',"
            ZSql = ZSql + " ImpoLetra = " + "'" + ZImpoLetra + "',"
            ZSql = ZSql + " VtoLetra = " + "'" + ZVtoLetra + "',"
            ZSql = ZSql + " VtoLetraII = " + "'" + ZVtoLetraII + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

        End If
        
    End If

    Call Conecta_Empresa

    PrgOrdenComplementoConsulta.Hide
    Unload Me
    Rem PrgArti.Show
    
End Sub

Private Sub Busca_Empresa()

    XEmpresa = Wempresa
    EmpresaOrden = 0
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then

        For Va = 1 To 7
    
            Select Case Va
                Case 1
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 5
                    Wempresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 6
                    Wempresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 7
                    Wempresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            spOrden = "ListaOrden " + "'" + Orden.Text + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
                If WTipoOrden = 1 Then
                    EmpresaOrden = Val(Wempresa)
                End If
                rstOrden.Close
            End If
        
        Next Va
        
            Else
        
        For Va = 1 To 4
    
            Select Case Va
                Case 1
                    Wempresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    Wempresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    Wempresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            spOrden = "ListaOrden " + "'" + Orden.Text + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
                If WTipoOrden = 1 Then
                    EmpresaOrden = Val(Wempresa)
                End If
                rstOrden.Close
            End If
        
        Next Va
        
    End If
    
    Call Conecta_Empresa
    
End Sub

