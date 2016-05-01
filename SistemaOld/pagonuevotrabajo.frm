VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form PrgPagoNuevoTrabajo 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Pagos a Proveedores"
   ClientHeight    =   7860
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7860
   ScaleWidth      =   11880
   Begin VB.Frame IngreCarpeta 
      Caption         =   "Ingreso de Importes para Carpetas"
      Height          =   2895
      Left            =   8760
      TabIndex        =   42
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton GrabaCarpeta 
         Caption         =   "Confirma"
         Height          =   375
         Left            =   840
         TabIndex        =   49
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Carpeta4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   47
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Carpeta3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   46
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Carpeta2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   45
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Carpeta1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   44
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Carpeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   43
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   840
         TabIndex        =   48
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ComboBox TipoPago 
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
      Left            =   120
      TabIndex        =   79
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox RetIbCiudad 
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
      Left            =   4800
      TabIndex        =   77
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Busqueda 
      Height          =   255
      Left            =   240
      TabIndex        =   76
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"pagonuevotrabajo.frx":0000
   End
   Begin VB.TextBox Lectora 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   75
      Top             =   3240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Frame IngreCuenta 
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
      Height          =   855
      Left            =   3360
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
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
         Left            =   1560
         TabIndex        =   28
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox TotRet 
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
      Left            =   10560
      TabIndex        =   73
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   12
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   5640
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   11
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   5400
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   10
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   5520
      Width           =   375
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
      Left            =   4920
      TabIndex        =   67
      Top             =   5040
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   4200
      TabIndex        =   66
      Top             =   4680
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
      Left            =   4200
      TabIndex        =   65
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox WTituloVector 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   5040
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3495
      Left            =   0
      TabIndex        =   56
      Top             =   3360
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.TextBox RetIva 
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
      Left            =   7800
      TabIndex        =   54
      Text            =   " "
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Paridad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   " "
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton CargaCarpeta 
      Caption         =   "Carpetas"
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
      Left            =   2040
      TabIndex        =   41
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox RetIb 
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
      Left            =   10560
      TabIndex        =   39
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Limpia1 
      Caption         =   "Limpia Renglon"
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
      Left            =   2040
      TabIndex        =   38
      Top             =   2640
      Width           =   1695
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
      Left            =   7200
      TabIndex        =   37
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton calcret 
      Caption         =   "Calc.Ret."
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
      Left            =   5040
      TabIndex        =   36
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Retencion 
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
      Left            =   7800
      TabIndex        =   35
      Text            =   " "
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Banco 
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
      Left            =   1680
      TabIndex        =   32
      Text            =   " "
      Top             =   1080
      Width           =   735
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
      Left            =   1680
      TabIndex        =   24
      Text            =   " "
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
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
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin Crystal.CrystalReport LISTADO 
      Left            =   8760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ordpago.rpt"
      WindowTitle     =   "Orden de Pago"
      CopiesToPrinter =   2
      WindowState     =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Orden de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   4455
      Begin VB.OptionButton Tipo6 
         Caption         =   "Aplic.Pago Impo."
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
         TabIndex        =   51
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Tipo5 
         Caption         =   "Ch. Rechazados"
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
         TabIndex        =   30
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Tipo4 
         Caption         =   "Transferencias"
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
         TabIndex        =   29
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Tipo3 
         Caption         =   "Pagos Varios"
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
         TabIndex        =   25
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Pagos Cta.Cte."
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
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1680
      MaxLength       =   11
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   1335
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
      Height          =   1425
      Left            =   7920
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   0
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
      Height          =   2010
      ItemData        =   "pagonuevotrabajo.frx":0083
      Left            =   7200
      List            =   "pagonuevotrabajo.frx":008A
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
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
      Height          =   300
      Left            =   6120
      TabIndex        =   9
      Top             =   2280
      Width           =   975
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
      Height          =   300
      Left            =   5040
      TabIndex        =   3
      Top             =   2280
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
      Height          =   300
      Left            =   6120
      TabIndex        =   8
      Top             =   1920
      Width           =   975
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
      Height          =   300
      Left            =   720
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Grabar"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4200
      TabIndex        =   68
      Top             =   4320
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
   Begin MSMask.MaskEdBox FechaParidad 
      Height          =   285
      Left            =   4680
      TabIndex        =   80
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   65535
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
   Begin VB.Label Label14 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Paridad"
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
      Left            =   4680
      TabIndex        =   81
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Ib Ciudad"
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
      Left            =   3120
      TabIndex        =   78
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Retenc."
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
      Left            =   9120
      TabIndex        =   74
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Iva"
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
      Left            =   6120
      TabIndex        =   55
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paridad"
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
      Left            =   6120
      TabIndex        =   53
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Dife 
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
      Left            =   10440
      TabIndex        =   50
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret.Ing.Brutos"
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
      Left            =   9120
      TabIndex        =   40
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ret. Ganancia"
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
      Left            =   6120
      TabIndex        =   34
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label DesBanco 
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
      TabIndex        =   33
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label bjm 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   360
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
      Left            =   10440
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Debitos 
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
      Left            =   2760
      TabIndex        =   19
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"pagonuevotrabajo.frx":0098
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
      Left            =   4200
      TabIndex        =   18
      Top             =   7080
      Width           =   6135
   End
   Begin VB.Label DesProveedor 
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
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   3975
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Orden de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "PrgPagoNuevoTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Debito As Double
Private Credito As Double
Private WImpresion(12, 10) As String
Private WImpre2(12, 10) As String
Private WDebito(12, 2) As String
Private WCredito(12, 4) As String
Private WCuenta(12, 2) As String
Private WCuentaBco As String
Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WSaldoUs As Double
Private WRetencion As Double
Private WRetIb As Double
Private WRetIbCiudad As Double
Private WRetIva As Double
Private WSumaIva As Double
Private WDife As Double
Private WCuatri  As String
Private WEmpNombre As String
Private WEmpDirecion As String
Private WEmpLocalidad As String
Private WEmpCuit As String
Private WPrvDireccion As String
Private WPrvCuit As String
Private WPrvIb As String
Private WLeyenda(10) As String
Private WTipo As String
Private WTipoprv As Single
Private WTipoiva As Single
Private WTipoIb As Single
Private WNeto As Double
Private WAnticipo As Double
Private WBruto As Double
Private WIva As Double
Private WRetenido As Double
Private WFecha As String
Private XNeto As Double
Private XBruto As Double
Private XIva As Double
Private XTBase As Double
Private XImpor As Double
Private WParametro(0 To 10) As Double
Private WTasa1(10) As Double
Private WAuxi As Double
Private WAuxi1 As Double
Private Total As Double
Private WRete1 As Double
Private WRete2 As Double
Private ZZBase As Double

Dim ZZClave As String
Dim ZZCorte As String
Dim ZZOrden As String
Dim ZZRenglon As String
Dim ZZProveedor As String
Dim ZZFecha As String
Dim ZZImporteCheque As String
Dim ZZNumeroCheque As String
Dim ZZFechaCheque As String
Dim ZZBancoCheque As String
Dim ZZCuit As String

Dim rstCtaCtePrv As Recordset
Dim spCtaCtePrv As String
Dim rstIvaComp As Recordset
Dim spIvaComp As String
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstPagos As Recordset
Dim spPagos As String
Dim rstPagosResumen As Recordset
Dim spPagosResumen As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim rstRetencion As Recordset
Dim spRetencion As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstNumero As Recordset
Dim spNumero As String
Dim XParam As String
Dim WProceso As Integer
Dim WCerti As String
Dim WCerificado As Integer
Dim ImpreCopia(10) As String
Dim WRete As Double
Dim WImpoRetenido As Double
Dim XImpre1 As String
Dim XImpre2 As String
Dim XImpre3 As String
Dim XImpre4 As String
Dim WImpre4 As Double
Dim WImpo1 As Double
Dim WImpo2 As Double
Dim WCertificadoGan As Integer
Dim WCertificadoIb As Integer
Dim WCertificadoIbCiudad As Integer
Dim WCertificadoIva As Integer
Dim Deuda(1000, 10) As String
Dim XNroInterno As String
Dim WTipoDife As String
Dim WLetraDife As String
Dim WPuntoDife As String
Dim WNumeroDife As String
Dim WNetoDife As Double
Dim WIvaDife As Double
Dim RenglonDife As Integer
Dim ParidadTotal As Double
Dim ZFecha As String
Dim ZTipo2 As String
Dim WPorceIb As Double
Dim WPorceIbCaba As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String
Dim WControlII As String
Dim CargaEmpresa(10, 2) As String
Dim ZCarpeta(10) As String
Dim XEmpresa As String

Dim ZEntraI(5000, 3) As String
Dim ZEntraII(5000, 5) As String

Private Sub Suma_Datos()

    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    ZZBase = 0
    
    For iRow = 1 To 12
        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
            If Tipo1.Value = True Then
            
                ZZZEntra = "N"
                ZZZLetra = WVector1.TextMatrix(iRow, 2)
                ZZZTipo = WVector1.TextMatrix(iRow, 1)
                ZZZPunto = WVector1.TextMatrix(iRow, 3)
                ZZZNumero = WVector1.TextMatrix(iRow, 4)
                If Trim(ZZZLetra) = "" And Val(ZZZTipo) = 0 And Val(ZZZPunto) = 0 And Val(ZZZNumero) = 0 Then
                
                    ZZZEntra = "S"
                        
                        Else
                
                    WLetra = ZZZLetra
                    WTipo = ZZZTipo
                    WPunto = ZZZPunto
                    WNumero = ZZZNumero
                    
                    ZRechazado = 0
                    ZNroInterno = "0"
                    
                    
                    ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM CtaCtePrv"
                    ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClaveCtaCtePrv + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCtaCtePrv.RecordCount > 0 Then
                        ZNroInterno = Str$(rstCtaCtePrv!NroInterno)
                        rstCtaCtePrv.Close
                    End If
                    
                    spIvaComp = "Consultaivacomp " + "'" + ZNroInterno + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
                        rstIvaComp.Close
                    End If
                    
                    If ZRechazado = 1 Then
                        ZZZEntra = "S"
                    End If
                    
                End If
                
                If Val(WVector1.TextMatrix(iRow, 1)) <> 0 Or ZZZEntra = "S" Then
                    Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(iRow, 5)))
                End If
                
                If ZZZEntra = "N" Then
                    ZZBase = ZZBase + Val(WVector1.TextMatrix(iRow, 5))
                End If
                
                    Else
                    
                Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(iRow, 5)))
                ZZBase = ZZBase + Val(WVector1.TextMatrix(iRow, 5))
                
            End If
        End If
        
        If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(WVector1.TextMatrix(iRow, 12)))
        End If
        
    Next iRow
    
    Existe = "N"
    
    ClavePagos = Orden.Text + "01"
    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Existe = "S"
        rstPagos.Close
    End If
    
    If Existe <> "S" Then
        Call calcret_Click
        Call CalcRetIb
    End If
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text) + Val(RetIb.Text) + Val(RetIbCiudad.Text) + Val(RetIva.Text))
    
    WDife = Val(Debitos.Caption) - Val(Creditos.Caption)
    Dife.Caption = Str$(WDife)
    
    Debitos.Caption = Pusing("#,###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("#,###,###.##", Creditos.Caption)
    Dife.Caption = Pusing("#,###,###.##", Dife.Caption)
    
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        ClavePagos = Orden.Text + Auxi1
    
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
            Select Case Val(rstPagos!Tiporeg)
                Case 1
                    Debito = Debito + 1
                    WVector1.TextMatrix(Debito, 1) = rstPagos!Tipo1
                    WVector1.TextMatrix(Debito, 2) = rstPagos!Letra1
                    WVector1.TextMatrix(Debito, 3) = rstPagos!Punto1
                    WVector1.TextMatrix(Debito, 4) = rstPagos!Numero1
                    WVector1.TextMatrix(Debito, 5) = Str$(rstPagos!Importe1)
                    Rem WVector1.TextMatrix(Debito, 5) = Pusing("#,###,###.##", WVector1.TextMatrix(Debito, 5))
                    WVector1.TextMatrix(Debito, 6) = rstPagos!Observaciones2
                    ZCta = IIf(IsNull(rstPagos!Cuenta), "", rstPagos!Cuenta)
                    WCuenta(Debito, 1) = ZCta
                    
                Case 2
                    Credito = Credito + 1
                    WVector1.TextMatrix(Credito, 7) = rstPagos!Tipo2
                    WVector1.TextMatrix(Credito, 8) = rstPagos!Numero2
                    WVector1.TextMatrix(Credito, 9) = rstPagos!Fecha2
                    WVector1.TextMatrix(Credito, 10) = rstPagos!Banco2
                    If rstPagos!Observaciones2 <> "" Then
                        WVector1.TextMatrix(Credito, 11) = rstPagos!Observaciones2
                    End If
                    WVector1.TextMatrix(Credito, 12) = Str$(rstPagos!Importe2)
                    WVector1.TextMatrix(Credito, 12) = Pusing("#,###,###.##", WVector1.TextMatrix(Credito, 12))
                    ZCta = IIf(IsNull(rstPagos!Cuenta), "", rstPagos!Cuenta)
                    WCuenta(Credito, 2) = ZCta
                Case Else
            End Select
            rstPagos.Close
                Else
            Exit Do
        End If
    Loop
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
    Rem Retganancias.text = PUsing("#,###,###.##", Retganancias.text)
End Sub

Sub Imprime_Datos()

    If Val(Banco.Text) <> 0 Then
        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBanco.Caption = rstBanco!Nombre
            rstBanco.Close
        End If
    End If

    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        Proveedor.Text = RstProveedor!Proveedor
        DesProveedor.Caption = RstProveedor!Nombre
        WPrvDireccion = RstProveedor!Direccion
        WPrvCuit = RstProveedor!Cuit
        WPrvIb = RstProveedor!NroIb
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WTipoiva = Val(RstProveedor!Iva)
        WTipoIb = RstProveedor!CodIb
        RstProveedor.Close
        Call Format_datos
    End If
    
End Sub

Private Sub cmdAdd_Click()

    If Fecha.Text <> "" Then
    
    If Proveedor.Text <> "" Or Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
    
    If Tipo4.Value = True And Val(Banco.Text) = 0 Then
        m$ = "No se ha informado el banco al cual se va a realizar al transferencia"
        A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
        Exit Sub
    End If
    
    If Tipo1.Value = False And Tipo2.Value = False Then
        If Val(Proveedor.Text) <> 0 Then
            m$ = "Solo se puede informar proveedor en las ordenes de pago de Pagos o Anticipos de Proveedores"
            A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
            Exit Sub
        End If
    End If
    
    ZProvincia = 0
    ZTipoProv = 0
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        ZProvincia = RstProveedor!Provincia
        ZTipoProv = IIf(IsNull(RstProveedor!TipoProv), "0", RstProveedor!TipoProv)
        RstProveedor.Close
    End If
    
   If Proveedor.Text = "10167878480" Or Proveedor.Text = "10000000100" Or Proveedor.Text = "10071081483" Or (Val(ZProvincia) = 24 And ZTipoProv = 1) Then
       If Val(Carpeta.Text) = 0 And Val(Carpeta1.Text) = 0 And Val(Carpeta2.Text) = 0 And Val(Carpeta3.Text) = 0 And Val(Carpeta4.Text) = 0 Then
            m$ = "Se debe informar el numero de carpeta correspondiente"
            A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
           Exit Sub
        End If
    End If
    
    If Val(Proveedor.Text) <> 0 Then
        If Val(Creditos.Caption) > 1000 Then
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                ZEmbargo = IIf(IsNull(RstProveedor!Embargo), "", RstProveedor!Embargo)
                RstProveedor.Close
                If ZEmbargo = "S" Then
                    T$ = "Emision de Ordenes de Pago"
                    m$ = "El proveedor tiene embargos por parte de Arba" + Chr$(13) + "Desea cancelar la emision de la orden de pago"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    Auxi1 = Orden.Text
    Call Ceros(Auxi1, 6)
    Orden.Text = Auxi1
    
    Existe = "N"
    
    ClavePagos = Orden.Text + "01"
    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Existe = "S"
        rstPagos.Close
    End If
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        If Val(Debitos.Caption) <> 0 Then
            Debito = Val(Debitos.Caption)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        If Debito = Credito Then
        
            If Val(Orden.Text) = 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select Max(Orden) as [OrdenMayor]"
                ZSql = ZSql + " FROM Pagos"
                spPagos = ZSql
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                If rstPagos.RecordCount > 0 Then
                    rstPagos.MoveLast
                    ZUltimo = IIf(IsNull(rstPagos!OrdenMayor), "0", rstPagos!OrdenMayor)
                    WOrden = Mid$(Str$(ZUltimo + 1), 2, 8)
                    rstPagos.Close
                        Else
                    WOrden = "1"
                End If
                Auxi1 = WOrden
                Call Ceros(Auxi1, 6)
                Orden.Text = Auxi1
                
                Rem spPagos = "ListaPagosNumero"
                Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                Rem If rstPagos.RecordCount > 0 Then
                Rem    With rstPagos
                Rem         .MoveLast
                Rem         Orden.Text = rstPagos!Orden + 1
                Rem         Auxi1 = Orden.Text
                Rem         Call Ceros(Auxi1, 6)
                Rem         Orden.Text = Auxi1
                Rem     End With
                Rem     rstPagos.Close
                Rem End If
                
            End If
            
            For iRow = 1 To 12
            
                If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
                    If Tipo3.Value = True Then
                        If WCuenta(iRow, 1) = "" Then
                            m$ = "No se ha imputado correctamente el concepto del pago"
                            A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos Varios")
                            Exit Sub
                        End If
                    End If
                End If
                
                ZTipo = WVector1.TextMatrix(iRow, 7)
                
                If Val(ZTipo) = 2 Then
                    ZFecha = WVector1.TextMatrix(iRow, 9)
                    Call Valida_fecha1(ZFecha, Auxi)
                    If Auxi = "S" And Len(ZFecha) = 10 Then
                        ZOrdFecha1 = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZOrdFecha2 = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
                        If ZOrdFecha2 < ZOrdFecha1 Then
                            m$ = "La Fecha de los valores informados no puede ser menor a la fecha de emision de la orden de pago"
                            A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos")
                            Exit Sub
                        End If
                            Else
                        m$ = "La Fecha de los valores informados es incorrecta"
                        A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos")
                        Exit Sub
                    End If
                End If
                
                If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Then
                    If Val(WVector1.TextMatrix(iRow, 7)) = 0 Then
                        m$ = "No se han informado correctamente los valores a entregar"
                        A% = MsgBox(m$, 0, "Emision de Ordenes de Pagos Varios")
                        Exit Sub
                    End If
                End If
                
            Next iRow
            
            WCertificadoGan = 0
            WCertificadoIb = 0
            WCertificadoIbCiudad = 0
            WCertificadoIva = 0
            
            If Val(Retencion.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "91" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoGan = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "91" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
            End If
            
            If Val(RetIb.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "92" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoIb = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "92" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
            
            End If
            
            If Val(RetIbCiudad.Text) <> 0 Then
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "94" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoIbCiudad = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "94" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
            End If
            
            If Val(RetIva.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Numero"
                ZSql = ZSql + " Where Numero.Codigo = " + "'" + "93" + "'"
                spNumero = ZSql
                Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                If rstNumero.RecordCount > 0 Then
                    WCertificadoIva = rstNumero!Numero + 1
                    rstNumero.Close
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Numero SET "
                    ZSql = ZSql + " Numero = Numero + 1"
                    ZSql = ZSql + " Where Codigo = " + "'" + "93" + "'"
                    spNumero = ZSql
                    Set rstNumero = db.OpenRecordset(spNumero, dbOpenSnapshot, dbSQLPassThrough)
                End If
                 
            End If
            
            
            
            Renglon = 0
            For iRow = 1 To 12
            
                WRow = iRow
                If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
                    
                    XNumero1 = WVector1.TextMatrix(iRow, 4)
                    If XNumero1 = "99999999" Then
                    
                        WTipoDife = Left$(WVector1.TextMatrix(iRow, 1), 2)
                        WLetraDife = Left$(WVector1.TextMatrix(iRow, 2), 1)
                        WPuntoDife = Left$(WVector1.TextMatrix(iRow, 3), 4)
                        
                        Select Case WLetraDife
                            Case "A"
                                WNetoDife = Val(WVector1.TextMatrix(iRow, 5)) / 1.21
                                Call Redondeo(WNetoDife)
                                WIvaDife = Val(WVector1.TextMatrix(iRow, 5)) - WNetoDife
                                Call Redondeo(WIvaDife)
                            Case Else
                                WNetoDife = Val(WVector1.TextMatrix(iRow, 5))
                                Call Redondeo(WNetoDife)
                                WIvaDife = 0
                        End Select
                        Call Alta_Dife
                        WVector1.TextMatrix(iRow, 4) = WNumeroDife
                    End If
                    
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    XOrden = Orden.Text
                    XRenglon = Auxi1
                    XProveedor = Proveedor.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Str$(Debito)
                    XRetencion = Retencion.Text
                    XRetotra = RetIb.Text
                    XRetIbCiudad = RetIbCiudad.Text
                    XRetIva = RetIva.Text
                    XObservaciones = Observaciones.Text
                    XCuenta = ""
                    If Tipo1.Value = True Then
                        XTipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        XTipoOrd = "2"
                    End If
                    If Tipo3.Value = True Then
                        XTipoOrd = "3"
                        XCuenta = WCuenta(iRow, 1)
                    End If
                    If Tipo4.Value = True Then
                        XTipoOrd = "4"
                        spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
                        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                        If rstBanco.RecordCount > 0 Then
                            XCuenta = rstBanco!Cuenta
                            rstBanco.Close
                                Else
                            XCuenta = "999999"
                        End If
                    End If
                    If Tipo5.Value = True Then
                        XTipoOrd = "5"
                        XCuenta = "111"
                    End If
                    If Tipo6.Value = True Then
                        XTipoOrd = "6"
                    End If
                    
                    ZZZValida = "S"
                    ZZZLetra = WVector1.TextMatrix(iRow, 2)
                    ZZZTipo = WVector1.TextMatrix(iRow, 1)
                    ZZZPunto = WVector1.TextMatrix(iRow, 3)
                    ZZZNumero = WVector1.TextMatrix(iRow, 4)
                    If Trim(ZZZLetra) = "" And Val(ZZZTipo) = 0 And Val(ZZZPunto) = 0 And Val(ZZZNumero) = 0 Then
                        XCuenta = WCuenta(iRow, 1)
                    End If
                    
                    XTiporeg = "1"
                    
                    XTipo1 = Left$(WVector1.TextMatrix(iRow, 1), 2)
                    XLetra1 = Left$(WVector1.TextMatrix(iRow, 2), 1)
                    XPunto1 = Left$(WVector1.TextMatrix(iRow, 3), 4)
                    XNumero1 = Left$(WVector1.TextMatrix(iRow, 4), 8)
                    XImporte1 = WVector1.TextMatrix(iRow, 5)
                    XObservaciones2 = Left$(WVector1.TextMatrix(iRow, 6), 30)
                    
                    XTipo2 = ""
                    XNumero2 = ""
                    XFecha2 = ""
                    XFechaOrd2 = ""
                    If Tipo4.Value = True Then
                        XBanco2 = Banco.Text
                            Else
                        XBanco2 = ""
                    End If
                    XImporte2 = ""
                    XEmpresa = "1"
                    XClave = XOrden + XRenglon
                    XRetganancias = ""
                    XConcepto = ""
                    XConcecionaria = ""
                    XImpolist = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                    spPagos = "AltaPagos " + XParam
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ClaveRecibo = ""
                    XCuit = ""
                    ImporteCheque = ""
                    NumeroCheque = ""
                    FechaCheque = ""
                    BancoCheque = ""
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Pagos SET "
                    ZSql = ZSql + " ClaveRecibo = " + "'" + ClaveRecibo + "',"
                    ZSql = ZSql + " RetIbCiudad = " + "'" + XRetIbCiudad + "',"
                    ZSql = ZSql + " ImporteCheque = " + "'" + ImporteCheque + "',"
                    ZSql = ZSql + " NumeroCheque = " + "'" + NumeroCheque + "',"
                    ZSql = ZSql + " FechaCheque = " + "'" + FechaCheque + "',"
                    ZSql = ZSql + " BancoCheque = " + "'" + BancoCheque + "',"
                    ZSql = ZSql + " Cuit = " + "'" + XCuit + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + XClave + "'"
                    spPagos = ZSql
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    WLetra = XLetra1
                    WTipo = XTipo1
                    WPunto = XPunto1
                    WNumero = XNumero1
                    WImporte = XImporte1
                    
                    If Tipo1.Value = True Then
                        ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                        spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                        Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        If RstCtaPrv.RecordCount > 0 Then
                            XSaldo = Str$(RstCtaPrv!Saldo - Val(WImporte))
                            XParam = "'" + ClaveCtaprv + "','" _
                                         + XSaldo + "'"
                            spCtaprv = "ActualizaCtaprvSaldo " + XParam
                            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                    
                End If
                
                If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Or Left$(WVector1.TextMatrix(iRow, 7), 2) <> "" Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    XOrden = Orden.Text
                    XRenglon = Auxi1
                    XProveedor = Proveedor.Text
                    XFecha = Fecha.Text
                    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XImporte = Str$(Debito)
                    XRetencion = Retencion.Text
                    XRetotra = RetIb.Text
                    XRetIbCiudad = RetIbCiudad.Text
                    XRetIva = RetIva.Text
                    XObservaciones = Observaciones.Text
                    If Tipo1.Value = True Then
                        XTipoOrd = "1"
                    End If
                    If Tipo2.Value = True Then
                        XTipoOrd = "2"
                    End If
                    If Tipo3.Value = True Then
                        XTipoOrd = "3"
                    End If
                    If Tipo4.Value = True Then
                        XTipoOrd = "4"
                    End If
                    If Tipo5.Value = True Then
                        XTipoOrd = "5"
                    End If
                    If Tipo6.Value = True Then
                        XTipoOrd = "6"
                    End If
                    XTiporeg = "2"
                    XTipo1 = ""
                    XLetra1 = ""
                    XPunto1 = ""
                    XNumero1 = ""
                    XImporte1 = ""
                    
                    XTipo2 = Left$(WVector1.TextMatrix(iRow, 7), 2)
                    XNumero2 = Left$(WVector1.TextMatrix(iRow, 8), 8)
                    XFecha2 = Left$(WVector1.TextMatrix(iRow, 9), 10)
                    XFechaOrd2 = Right$(XFecha2, 4) + Mid$(XFecha2, 4, 2) + Left$(XFecha2, 2)
                    XBanco2 = WVector1.TextMatrix(iRow, 10)
                    XObservaciones2 = Left$(WVector1.TextMatrix(iRow, 11), 20)
                    XImporte2 = WVector1.TextMatrix(iRow, 12)
                    TipoRecibos = Left$(WVector1.TextMatrix(iRow, 13), 1)
                    ClaveRecibos = Mid$(WVector1.TextMatrix(iRow, 13), 2, 10)
                    Cuit = WVector1.TextMatrix(iRow, 14)
                    ClaveCtacte = WVector1.TextMatrix(iRow, 13)
                    
                    XEmpresa = "1"
                    XClave = XOrden + XRenglon
                    XRetganancias = ""
                    XConcepto = ""
                    XConcecionaria = ""
                    XImpolist = ""
                    XCuenta = ""
                    If Val(XTipo2) = 6 Then
                        XCuenta = WCuenta(iRow, 2)
                    End If
                    If Val(XTipo2) = 2 Then
                        ZTipo2 = XTipo2
                        Call Ceros(ZTipo2, 2)
                        XTipo2 = ZTipo2
                    End If
                    
                    XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                    spPagos = "AltaPagos " + XParam
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ClaveRecibo = ClaveRecibos
                    XCuit = Cuit
                    ImporteCheque = XImporte2
                    NumeroCheque = XNumero2
                    FechaCheque = XFecha2
                    BancoCheque = XObservaciones2
                    
                    ZClaveRecibo = Right$(ClaveRecibo, 8)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Pagos SET "
                    ZSql = ZSql + " ClaveRecibo = " + "'" + ZClaveRecibo + "',"
                    ZSql = ZSql + " RetIbCiudad = " + "'" + XRetIbCiudad + "',"
                    ZSql = ZSql + " ImporteCheque = " + "'" + ImporteCheque + "',"
                    ZSql = ZSql + " NumeroCheque = " + "'" + NumeroCheque + "',"
                    ZSql = ZSql + " FechaCheque = " + "'" + FechaCheque + "',"
                    ZSql = ZSql + " BancoCheque = " + "'" + BancoCheque + "',"
                    ZSql = ZSql + " Cuit = " + "'" + XCuit + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + XClave + "'"
                    spPagos = ZSql
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If Val(XTipo2) = 3 Then
                    
                        Rem spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibos + "'"
                        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstRecibos.RecordCount > 0 Then
                        Rem     XEstado2 = "X"
                        Rem     XDestino = ""
                        Rem     XParam = "'" + ClaveRecibos + "','" _
                        rem                  + XEstado2 + "','" _
                        rem                  + XDestino + "'"
                        Rem     spRecibos = "ActualizaRecibos " + XParam
                        Rem     Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        Rem End If
                        
                        Rem hay que verificar esto
                        
                        XEstado2 = "X"
                        XDestino = "O.P:" + Orden.Text
                        
                        If TipoRecibos = "1" Then
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Recibos SET "
                            ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
                            ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + ClaveRecibo + "'"
                            spRecibos = ZSql
                            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                                Else
                            ZSql = ""
                            ZSql = ZSql + "UPDATE RecibosProvi SET "
                            ZSql = ZSql + "Estado2 = " + "'" + "X" + "',"
                            ZSql = ZSql + "Destino = " + "'" + XObservaciones + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + ClaveRecibo + "'"
                            spRecibosProvi = ZSql
                            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                        
                        Rem ZSql = ""
                        Rem ZSql = ZSql + "UPDATE Recibos SET "
                        Rem ZSql = ZSql + " Estado2 = " + "'" + XEstado2 + "',"
                        Rem ZSql = ZSql + " Destino = " + "'" + XDestino + "'"
                        Rem ZSql = ZSql + " Where Numero2 = " + "'" + XNumero2 + "'"
                        Rem ZSql = ZSql + " and Importe2 = " + "'" + XImporte2 + "'"
                        
                        Rem spRecibos = ZSql
                        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Val(XTipo2) = 4 Then
                        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCte.RecordCount > 0 Then
                            XSaldo = ""
                            XSaldoUs = ""
                            XEstado = "1"
                            XDate = Date$
                            rstCtaCte.Close
                            XParam = "'" + ClaveCtacte + "','" _
                                         + XSaldo + "','" _
                                         + XSaldoUs + "','" _
                                         + XEstado + "','" _
                                         + XDate + "'"
                            spCtaCte = "ActualizaCtacte " + XParam
                            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                    
                End If
                
            Next iRow
            
            XParam = "'" + Orden.Text + "','" _
                    + Paridad.Text + "'"
            spPagos = "ModificaPagosParidad " + XParam
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                Claveven$ = "1"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WCtaProveedor = !CtaProveedores
                    WCtaEfectivo = !CtaEfectivo
                    WCtaCheques = !CtaCheque
                End If
            End With
        
            If Tipo1.Value = True Then
        
                WLetra = "A"
                WTipo = "04"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Proveedor.Text
        
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WProveedor, 6)
        
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount > 0 Then
            
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = ""
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "OP"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "'"
                    
                    spCtaprv = "ActualizaCtaCtePrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = ""
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "OP"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                    
                    spCtaprv = "AltaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
            End If
        
            If Tipo2.Value = True Then
        
                WLetra = "A"
                WTipo = "05"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Proveedor.Text
        
                Call Ceros(WNumero, 8)
                Rem Call Ceros(WProveedor, 6)
            
                ClaveCtaprv = WProveedor + WLetra + WTipo + WPunto + WNumero
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount > 0 Then
            
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = Str$(Debito * -1)
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "AN"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                            + XProveedor + "','" + XLetra + "','" _
                            + XTipo + "','" _
                            + XPunto + "','" + XNumero + "','" _
                            + XFecha + "','" _
                            + XEstado + "','" _
                            + Xvencimiento + "','" + XVencimiento1 + "','" _
                            + XTotal + "','" _
                            + XSaldo + "','" _
                            + XOrdFecha + "','" + XOrdVencimiento + "','" _
                            + XImpre + "','" + XEmpresa + "','" _
                            + XSaldolist + "','" _
                            + XNroInterno + "','" + Xlista + "','" _
                            + XAcumulado + "','" _
                            + XParidad + "','" _
                            + XPAgo + "'"
                        
                    spCtaprv = "ModificaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                        
                    XProveedor = Proveedor.Text
                    XLetra = WLetra
                    XTipo = WTipo
                    XPunto = WPunto
                    XNumero = WNumero
                    XFecha = Fecha.Text
                    XEstado = "1"
                    Xvencimiento = "  /  /    "
                    XVencimiento1 = "  /  /    "
                    XTotal = Str$(Debito * -1)
                    XSaldo = Str$(Debito * -1)
                    XClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    XOrdVencimiento = "00000000"
                    XImpre = "AN"
                    XEmpresa = "1"
                    XSaldolist = ""
                    XNroInterno = ""
                    Xlista = ""
                    XAcumulado = ""
                    XParidad = ""
                    XPAgo = ""
                    
                    XParam = "'" + XClave + "','" _
                        + XProveedor + "','" + XLetra + "','" _
                        + XTipo + "','" _
                        + XPunto + "','" + XNumero + "','" _
                        + XFecha + "','" _
                        + XEstado + "','" _
                        + Xvencimiento + "','" + XVencimiento1 + "','" _
                        + XTotal + "','" _
                        + XSaldo + "','" _
                        + XOrdFecha + "','" + XOrdVencimiento + "','" _
                        + XImpre + "','" + XEmpresa + "','" _
                        + XSaldolist + "','" _
                        + XNroInterno + "','" + Xlista + "','" _
                        + XAcumulado + "','" _
                        + XParidad + "','" _
                        + XPAgo + "'"
                    
                    spCtaprv = "AltaCtaPrv " + XParam
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
            End If
            
            If Tipo6.Value = True Then
            
                Renglon = Renglon + 1
                Auxi1 = Str$(Renglon)
                Call Ceros(Auxi1, 2)
                XOrden = Orden.Text
                XRenglon = Auxi1
                XProveedor = Proveedor.Text
                XFecha = Fecha.Text
                XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                XImporte = ""
                XRetencion = Retencion.Text
                XRetotra = RetIb.Text
                XRetIbCiudad = RetIbCiudad.Text
                XRetIva = RetIva.Text
                XObservaciones = Observaciones.Text
                XCuenta = ""
                XTipoOrd = "6"
                    
                XTiporeg = "1"
                XTipo1 = ""
                XLetra1 = ""
                XPunto1 = ""
                XNumero1 = ""
                XImporte1 = ""
                XObservaciones2 = "Aplicaicon de Pgos de Importacion"
                XTipo2 = ""
                XNumero2 = ""
                XFecha2 = ""
                XFechaOrd2 = ""
                XBanco2 = ""
                XImporte2 = ""
                XEmpresa = "1"
                XClave = XOrden + XRenglon
                XRetganancias = ""
                XConcepto = ""
                XConcecionaria = ""
                XImpolist = ""
                    
                XParam = "'" + XClave + "','" _
                            + XOrden + "','" + XRenglon + "','" _
                            + XProveedor + "','" _
                            + XFecha + "','" + XFechaOrd + "','" _
                            + XTipoOrd + "','" _
                            + XRetganancias + "','" _
                            + XRetIva + "','" + XRetotra + "','" _
                            + XRetencion + "','" _
                            + XTiporeg + "','" _
                            + XTipo1 + "','" + XLetra1 + "','" _
                            + XPunto1 + "','" + XNumero1 + "','" _
                            + XImporte1 + "','" _
                            + XTipo2 + "','" + XNumero2 + "','" _
                            + XFecha2 + "','" + XBanco2 + "','" _
                            + XImporte2 + "','" + XObservaciones2 + "','" _
                            + XEmpresa + "','" + XConcepto + "','" _
                            + XObservaciones + "','" _
                            + XImporte + "','" + XFechaOrd2 + "','" _
                            + XConcesionaria + "','" _
                            + XImpolist + "','" _
                            + XCuenta + "'"
                
                spPagos = "AltaPagos " + XParam
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                
                ClaveRecibo = ""
                XCuit = ""
                ImporteCheque = ""
                NumeroCheque = ""
                FechaCheque = ""
                BancoCheque = ""
                ZSql = ""
                ZSql = ZSql + "UPDATE Pagos SET "
                ZSql = ZSql + " ClaveRecibo = " + "'" + ClaveRecibo + "',"
                ZSql = ZSql + " RetIbCiudad = " + "'" + XRetIbCiudad + "',"
                ZSql = ZSql + " ImporteCheque = " + "'" + ImporteCheque + "',"
                ZSql = ZSql + " NumeroCheque = " + "'" + NumeroCheque + "',"
                ZSql = ZSql + " FechaCheque = " + "'" + FechaCheque + "',"
                ZSql = ZSql + " BancoCheque = " + "'" + BancoCheque + "',"
                ZSql = ZSql + " Cuit = " + "'" + XCuit + "'"
                ZSql = ZSql + " Where Clave = " + "'" + XClave + "'"
                spPagos = ZSql
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)

            End If
        
            If TipoPago.ListIndex = 0 Then
                ClaveRetencion = WFecha + Proveedor.Text
                spRetencion = "ConsultaRetencion " + "'" + ClaveRetencion + "'"
                Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
                If rstRetencion.RecordCount > 0 Then
                    XXNeto = Str$(rstRetencion!Neto + XNeto)
                    XXRetenido = Str$(rstRetencion!Retenido + Val(Retencion.Text))
                    XParam = "'" + ClaveRetencion + "','" + XXNeto + "','" _
                         + XXRetenido + "'"
                    rstRetencion.Close
                    spRetencion = "ActualizaRetencionPagos " + XParam
                    Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pagos SET "
            ZSql = ZSql + " Carpeta = " + "'" + Carpeta.Text + "',"
            ZSql = ZSql + " Carpeta1 = " + "'" + Carpeta1.Text + "',"
            ZSql = ZSql + " Carpeta2 = " + "'" + Carpeta2.Text + "',"
            ZSql = ZSql + " Carpeta3 = " + "'" + Carpeta3.Text + "',"
            ZSql = ZSql + " Carpeta4 = " + "'" + Carpeta4.Text + "'"
            ZSql = ZSql + " Where Orden = " + "'" + Orden.Text + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            
            ZCarpeta(1) = Carpeta.Text
            ZCarpeta(2) = Carpeta1.Text
            ZCarpeta(3) = Carpeta2.Text
            ZCarpeta(4) = Carpeta3.Text
            ZCarpeta(5) = Carpeta4.Text
            
            For Ciclo = 1 To 5
            
                If Val(ZCarpeta(Ciclo)) <> 0 Then
                
                    XEmpresa = WEmpresa
                    WEntra = "N"
        
                    Select Case Val(XEmpresa)
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
                    
                    For Cicla = 1 To ZHasta
                        If CargaEmpresa(Cicla, 1) <> "" Then
                
                            WEmpresa = CargaEmpresa(Cicla, 1)
                            txtOdbc = CargaEmpresa(Cicla, 2)
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Orden"
                            ZSql = ZSql + " Where Orden.Carpeta = " + "'" + ZCarpeta(Ciclo) + "'"
                            spOrden = ZSql
                            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                            If rstOrden.RecordCount > 0 Then
                                rstOrden.Close
                                WEntra = "S"
                            End If
                            
                            If WEntra = "S" Then
                            
                                If Proveedor.Text = "10167878480" Or Proveedor.Text = "10000000100" Or Proveedor.Text = "10071081483" Or Proveedor.Text = "10022098824" Then
                                    ZSql = ""
                                    ZSql = ZSql + "UPDATE Orden SET "
                                    ZSql = ZSql + " PagoDespacho = 1"
                                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + ZCarpeta(Ciclo) + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                        Else
                                    ZSql = ""
                                    ZSql = ZSql + "UPDATE Orden SET "
                                    ZSql = ZSql + " PagoLetra = 1"
                                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + ZCarpeta(Ciclo) + "'"
                                    spOrden = ZSql
                                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                            
                                Exit For
                                
                            End If
                    
                        End If
                    Next Cicla
    
                    Call Conecta_Empresa
                    
                End If
            
            Next
            
            
            
            
            
            Sql1 = "UPDATE Pagos SET "
            Sql2 = " CertificadoGan = " + "'" + Str$(WCertificadoGan) + "',"
            Sql3 = " CertificadoIb = " + "'" + Str$(WCertificadoIb) + "',"
            Sql4 = " CertificadoIbCiudad = " + "'" + Str$(WCertificadoIbCiudad) + "',"
            Sql5 = " CertificadoIva = " + "'" + Str$(WCertificadoIva) + "'"
            Sql6 = " Where Orden = " + "'" + Orden.Text + "'"
            spPagos = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
            With rstEmpresa
                .Index = "Empresa"
                .Seek "=", Val(WEmpresa)
                If .NoMatch = False Then
                    WAuxiliar = !Nombre
                End If
            End With
        
            Call IMPREORDEN
            If Val(Retencion.Text) <> 0 Then
                Call Impreret
            End If
            If Val(RetIb.Text) <> 0 Then
                Call Impreretib
            End If
            If Val(RetIbCiudad.Text) <> 0 Then
                Call ImpreretibCiudad
            End If
            If Val(RetIva.Text) <> 0 Then
                Call ImpreretIva
            End If

            Orden.SetFocus
            Call CmdLimpiar_Click
        End If
        
    End If
    
    End If
    
    End If
End Sub

Private Sub cmdDelete_Click()
    If Orden.Text <> "" Then
                
            Rem Borro los datos anteriores
            
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Orden.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow

    End If
    Proveedor.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector
    TipoPago.ListIndex = 0
    
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Tipo6.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    RetIb.Text = ""
    RetIbCiudad.Text = ""
    RetIva.Text = ""
    WTipoprv = 0
    ParidadTotal = 0
    Existe = "N"
    
    Carpeta.Text = ""
    Carpeta1.Text = ""
    Carpeta2.Text = ""
    Carpeta3.Text = ""
    Carpeta4.Text = ""
    
    spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambio.RecordCount > 0 Then
        Paridad.Text = Str$(rstCambio!Cambio)
        Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
        rstCambio.Close
    End If
    
    Orden.SetFocus
    Orden.Text = ""
    
    Rem spPagos = "ListaPagosNumero"
    Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPagos.RecordCount > 0 Then
    Rem     With rstPagos
    Rem         .MoveLast
    Rem         Orden.Text = rstPagos!Orden + 1
    Rem     End With
    Rem     rstPagos.Close
    Rem End If
    
    Pantalla.Visible = False
    Opcion.Visible = False
    IngreCuenta.Visible = False
    Erase WCuenta
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    
    With rstEmpresa
        .Close
    End With
    
    Orden.SetFocus
    PrgPagoNuevo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Command1_Click()
    Rem ClaveCtaprv = "0100000012"
    Rem spCtaprv = "BorrarCtaprv " + "'" + ClaveCtaprv + "'"
    Rem Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_ImprePago
    OPEN_FILE_ImpreRetIb
    OPEN_FILE_ImpreRetGan
End Sub

Private Sub Impresion_Click()

    Existe = "N"
    
    ClavePagos = Orden.Text + "01"
    spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Existe = "S"
        rstPagos.Close
    End If
    
    If Existe = "S" Then

        WOrden = Orden.Text
        Call CmdLimpiar_Click
        Orden.Text = WOrden

        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        Existe = "N"
        
        ClavePagos = Orden.Text + "01"
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
        
            Existe = "S"
            Proveedor.Text = rstPagos!Proveedor
            Fecha.Text = rstPagos!Fecha
            Retencion.Text = rstPagos!Retencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
            
            Rem nan WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
            
           Rem RetIbCiudad.Text = rstPagos!RetIbCiudad
            RetIbCiudad.Text = IIf(IsNull(rstPagos!RetIbCiudad), "", rstPagos!RetIbCiudad)
            
            RetIbCiudad.Text = Pusing("#,###,###.##", RetIbCiudad.Text)
            RetIva.Text = rstPagos!RetIva
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
            Tipo6.Value = False
            Select Case Val(rstPagos!TipoOrd)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case 3
                    Tipo3.Value = True
                Case 4
                    Tipo4.Value = True
                Case 5
                    Tipo5.Value = True
                Case 6
                    Tipo6.Value = True
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            Carpeta1.Text = IIf(IsNull(rstPagos!Carpeta1), "", rstPagos!Carpeta1)
            Carpeta2.Text = IIf(IsNull(rstPagos!Carpeta2), "", rstPagos!Carpeta2)
            Carpeta3.Text = IIf(IsNull(rstPagos!Carpeta3), "", rstPagos!Carpeta3)
            Carpeta4.Text = IIf(IsNull(rstPagos!Carpeta4), "", rstPagos!Carpeta4)
            
            WCertificadoGan = IIf(IsNull(rstPagos!certificadoGan), "0", rstPagos!certificadoGan)
            WCertificadoIb = IIf(IsNull(rstPagos!CertificadoIb), "0", rstPagos!CertificadoIb)
            WCertificadoIbCiudad = IIf(IsNull(rstPagos!CertificadoIbCiudad), "0", rstPagos!CertificadoIbCiudad)
            WCertificadoIva = IIf(IsNull(rstPagos!CertificadoIva), "0", rstPagos!CertificadoIva)
            
            rstPagos.Close
                
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            
            Call IMPREORDEN
            
            If Val(Retencion.Text) <> 0 Then
                WRetencion = Val(Retencion.Text)
                Call Impreret
            End If
            
            If Val(RetIb.Text) <> 0 Then
                WRetIb = Val(RetIb.Text)
                Call Impreretib
            End If
            
            If Val(RetIbCiudad.Text) <> 0 Then
                WRetIbCiudad = Val(RetIbCiudad.Text)
                Call ImpreretibCiudad
            End If
            
            If Val(RetIva.Text) <> 0 Then
                WRetIva = Val(RetIva.Text)
                Call ImpreretIva
            End If
            
        End If
    End If
    
End Sub

Private Sub Limpia1_Click()

    WTexto1.Text = ""
    WTexto2.Text = ""
        
    A = WVector1.Row
    B = WVector1.Col
        
    If B <= 6 Then
    
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
        WVector1.TextMatrix(WVector1.Row, 6) = ""
        
            Else
            
        WVector1.TextMatrix(WVector1.Row, 7) = ""
        WVector1.TextMatrix(WVector1.Row, 8) = ""
        WVector1.TextMatrix(WVector1.Row, 9) = ""
        WVector1.TextMatrix(WVector1.Row, 10) = ""
        WVector1.TextMatrix(WVector1.Row, 11) = ""
        WVector1.TextMatrix(WVector1.Row, 12) = ""
        
    End If
        
    Call Suma_Datos
        
    If B <= 6 Then
        WVector1.Row = A
        WVector1.Col = 1
            Else
        WVector1.Row = A
        WVector1.Col = 7
    End If
        
    Call StartEdit
        
End Sub

Private Sub Orden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        Existe = "N"
        
        ClavePagos = Orden.Text + "01"
        spPagos = "ConsultaPagos " + "'" + ClavePagos + "'"
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
        
            Existe = "S"
            Proveedor.Text = rstPagos!Proveedor
            Fecha.Text = rstPagos!Fecha
            Retencion.Text = rstPagos!Retencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            RetIb.Text = rstPagos!RetOtra
            RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
           Rem RetIbCiudad.Text = rstPagos!RetIbCiudad
           RetIbCiudad.Text = IIf(IsNull(rstPagos!RetIbCiudad), "0", rstPagos!RetIbCiudad)
           
           RetIbCiudad.Text = Pusing("#,###,###.##", RetIbCiudad.Text)
                  
            
            RetIva.Text = rstPagos!RetIva
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Tipo5.Value = False
            Tipo6.Value = False
            Select Case Val(rstPagos!TipoOrd)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case 3
                    Tipo3.Value = True
                Case 4
                    Banco.Text = rstPagos!Banco2
                    Tipo4.Value = True
                Case 5
                    Tipo5.Value = True
                Case 6
                    Tipo6.Value = True
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            
            Carpeta.Text = IIf(IsNull(rstPagos!Carpeta), "", rstPagos!Carpeta)
            Carpeta1.Text = IIf(IsNull(rstPagos!Carpeta1), "", rstPagos!Carpeta1)
            Carpeta2.Text = IIf(IsNull(rstPagos!Carpeta2), "", rstPagos!Carpeta2)
            Carpeta3.Text = IIf(IsNull(rstPagos!Carpeta3), "", rstPagos!Carpeta3)
            Carpeta4.Text = IIf(IsNull(rstPagos!Carpeta4), "", rstPagos!Carpeta4)
            
            WCertificadoGan = IIf(IsNull(rstPagos!certificadoGan), "0", rstPagos!certificadoGan)
            WCertificadoIb = IIf(IsNull(rstPagos!CertificadoIb), "0", rstPagos!CertificadoIb)
            WCertificadoIbCiudad = IIf(IsNull(rstPagos!CertificadoIbCiudad), "0", rstPagos!CertificadoIbCiudad)
            WCertificadoIva = IIf(IsNull(rstPagos!CertificadoIva), "0", rstPagos!CertificadoIva)
            
            Paridad.Text = IIf(IsNull(rstPagos!Paridad), "0", rstPagos!Paridad)
            Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
            
            rstPagos.Close
                
        End If
        
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
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
                Observaciones.SetFocus
                    Else
                Proveedor.SetFocus
            End If
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub FechaParidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = ""
        Call Valida_fecha1(FechaParidad.Text, Auxi)
        If Auxi = "S" Then
            spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambio.RecordCount > 0 Then
                Paridad.Text = Str$(rstCambio!Cambio)
                Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
                rstCambio.Close
            End If
        End If
    End If
End Sub


Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Proveedor.Text = RstProveedor!Proveedor
                DesProveedor.Caption = RstProveedor!Nombre
                WPrvDireccion = RstProveedor!Direccion
                WPrvCuit = RstProveedor!Cuit
                WPrvIb = RstProveedor!NroIb
                WTipoprv = Val(RstProveedor!Tipo) + 1
                WTipoiva = Val(RstProveedor!Iva)
                WTipoIb = RstProveedor!CodIb
                ZEmbargo = IIf(IsNull(RstProveedor!Embargo), "", RstProveedor!Embargo)
                If ZEmbargo = "S" Then
                    m$ = "El proveedor tiene embargos por parte de Arba" + Chr$(13) + "Verifique el monto correspondiente antes de realizar el pago"
                    A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                End If
                RstProveedor.Close
                Observaciones.SetFocus
                    Else
                Proveedor.Text = Proveedor.Text
                Proveedor.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo4.Value = True Then
            Banco.SetFocus
                Else
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Banco.Text) <> 0 Then
            spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                Banco.Text = rstBanco!Banco
                DesBanco.Caption = rstBanco!Nombre
                WCtabanco = rstBanco!Cuenta
                rstBanco.Close
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                    Else
                Banco.Text = Banco.Text
                Banco.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
    
     XRow = WVector1.Row
     XCol = WVector1.Col

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Cuenta Corrientes"
     Opcion.AddItem "Cheques terceros"
     Opcion.AddItem "Documentos"
     Opcion.AddItem "Cuentas Contables"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            Rem spProveedor = "ListaProveedoresordConsulta"
            Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            Rem If RstProveedor.RecordCount > 0 Then
            Rem
            Rem     With RstProveedor
            Rem         .MoveFirst
            Rem         Do
            Rem             If .EOF = False Then
            Rem                 Auxi$ = Mascara("###########", Str$(RstProveedor!Proveedor))
            Rem                 Call Ceros(Auxi, 11)
            Rem                 IngresaItem = Auxi + "      " + RstProveedor!Nombre
            Rem                 Pantalla.AddItem IngresaItem
            Rem                 IngresaItem = RstProveedor!Proveedor
            Rem                 WIndice.AddItem IngresaItem
            Rem                 .MoveNext
            Rem                     Else
            Rem                 Exit Do
            Rem             End If
            Rem         Loop
            Rem     End With
            Rem     RstProveedor.Close
            Rem
            Rem End If
            
            Ayuda.SetFocus
            
        Case 1
            Erase Deuda
            EntraDeuda = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCtePrv"
            ZSql = ZSql + " Where CtaCtePrv.Proveedor = " + "'" + Proveedor.Text + "'"
            ZSql = ZSql + " and CtaCtePrv.Saldo <> 0"
            ZSql = ZSql + " Order by CtaCteprv.Proveedor, CtaCteprv.OrdFecha"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCtePrv.RecordCount > 0 Then
            Rem XParam = "'" + Proveedor.Text + "','" _
            rem             + Proveedor.Text + "'"
            Rem spCtaprv = "ListaCtaPrvDesdeHasta " + XParam
            Rem Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            Rem  If RstCtaPrv.RecordCount > 0 Then
            
                With rstCtaCtePrv
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Proveedor.Text = rstCtaCtePrv!Proveedor Then
                                WAuxi1 = rstCtaCtePrv!Saldo
                                Call Redondeo(WAuxi1)
                                If WAuxi1 <> 0 Then
                                    EntraDeuda = EntraDeuda + 1
                                    Deuda(EntraDeuda, 1) = !NroInterno
                                    Deuda(EntraDeuda, 2) = !Total
                                    Deuda(EntraDeuda, 3) = !Saldo
                                    Deuda(EntraDeuda, 4) = !Impre
                                    Deuda(EntraDeuda, 5) = !Letra
                                    Deuda(EntraDeuda, 6) = !Punto
                                    Deuda(EntraDeuda, 7) = !Numero
                                    Deuda(EntraDeuda, 8) = !Fecha
                                    Deuda(EntraDeuda, 9) = !Clave
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtaCtePrv.Close
                
            End If
            
            For Ciclo = 1 To EntraDeuda
            
                XNroInterno = Deuda(Ciclo, 1)
                XTotal = Deuda(Ciclo, 2)
                XSaldo = Deuda(Ciclo, 3)
                XImpre = Deuda(Ciclo, 4)
                XLetra = Deuda(Ciclo, 5)
                XPunto = Deuda(Ciclo, 6)
                XNumero = Deuda(Ciclo, 7)
                XFecha = Deuda(Ciclo, 8)
                XClave = Deuda(Ciclo, 9)

                spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                    XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                    rstIvaComp.Close
                End If
                                
                ParidadTotal = 0
                If XPAgo <> 2 Then
                    WSaldo = XSaldo
                    WSaldoUs = 0
                    Call Redondeo(WSaldo)
                        Else
                    spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
                    Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCambio.RecordCount > 0 Then
                        ParidadTotal = rstCambio!Cambio
                        rstCambio.Close
                    End If
                                    
                    WSaldo = XSaldo
                    WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                    Call Redondeo(WSaldo)
                    Call Redondeo(WSaldoUs)
                End If
                
                Auxi$ = Str$(WSaldo)
                Auxi$ = Mascara("#,###,###.##", Auxi$)
                If WSaldoUs <> 0 Then
                    Auxi1$ = Str$(WSaldoUs)
                    Auxi1$ = Mascara("#,###,###.##", Auxi1$)
                        Else
                    Auxi1$ = ""
                End If
                IngresaItem = XImpre + " " + XLetra + " " + XPunto + " " + XNumero + " " + XFecha + " " + Auxi$ + " " + Auxi1$
                Pantalla.AddItem IngresaItem
                IngresaItem = XClave
                WIndice.AddItem IngresaItem
                
            Next Ciclo
           
        Case 2
            ZSql = ""
            ZSql = ZSql + "UPDATE RecibosProvi SET "
            ZSql = ZSql + " ReciboDefinitivo = 0"
            ZSql = ZSql + " Where ReciboDefinitivo is null"
            spRecibosProvi = ZSql
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        
            Erase ZEntraI
            Erase ZEntraII
            
            ZLugarI = 0
            ZLugarII = 0
        
            ZSql = ""
            ZSql = ZSql + "Select Recibos.Tiporeg, Recibos.Estado2, Recibos.Tipo2, Recibos.TipoReg, Recibos.Importe2, Recibos.Numero2, Recibos.Fecha2, Recibos.Banco2, Recibos.Clave, Recibos.FechaOrd2"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.TipoReg = '2'"
            ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
            ZSql = ZSql + " and Recibos.Tipo2 = '02'"
            ZSql = ZSql + " Order by Recibos.FechaOrd2, Recibos.Numero2"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
            Rem spRecibos = "ListaRecibosNroCheque"
            Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstRecibos.RecordCount Then
            
                With rstRecibos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstRecibos!Tiporeg) = 2 Then
                                If Val(rstRecibos!Tipo2) = 2 And rstRecibos!Estado2 <> "X" Then
                                
                                    ZLugarI = ZLugarI + 1
                                    Auxi$ = Str$(rstRecibos!Importe2)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibos!Numero2))
                                    Call Ceros(Numero, 6)
                                    IngresaItem = Numero + "  " + rstRecibos!Fecha2 + "  " + Auxi$ + "  " + rstRecibos!Banco2
                                    
                                    WOrdFecha2 = IIf(IsNull(rstRecibos!FechaOrd2), "", rstRecibos!FechaOrd2)
                                
                                    ZEntraI(ZLugarI, 1) = IngresaItem
                                    ZEntraI(ZLugarI, 2) = "1" + rstRecibos!Clave
                                    ZEntraI(ZLugarI, 3) = WOrdFecha2
                                    
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
                
            End If
            
            Sql1 = "Select RecibosProvi.Tiporeg, RecibosProvi.Estado2, RecibosProvi.Tipo2, RecibosProvi.TipoReg, RecibosProvi.Importe2, RecibosProvi.Numero2, RecibosProvi.Fecha2, RecibosProvi.Banco2, RecibosProvi.Clave, RecibosProvi.FechaOrd2, RecibosProvi.ReciboDefinitivo"
            Sql2 = " FROM RecibosProvi"
            Sql3 = " Where RecibosProvi.TipoReg = " + "'" + "2" + "'"
            Sql4 = " and RecibosProvi.Estado2 = " + "'" + "P" + "'"
            Sql5 = " and RecibosProvi.ReciboDefinitivo = " + "'" + "0" + "'"
            Sql6 = " Order by FechaOrd2, Numero2"
            spRecibosProvi = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibosProvi.RecordCount > 0 Then
            
                With rstRecibosProvi
                    .MoveFirst
                    Do
                        If .EOF = False Then
                    
                            WTiporeg = IIf(IsNull(rstRecibosProvi!Tiporeg), "", rstRecibosProvi!Tiporeg)
                            WTipo2 = IIf(IsNull(rstRecibosProvi!Tipo2), "", rstRecibosProvi!Tipo2)
                            WEstado2 = IIf(IsNull(rstRecibosProvi!Estado2), "", rstRecibosProvi!Estado2)
                            WDefinitivo = IIf(IsNull(rstRecibosProvi!ReciboDefinitivo), "0", rstRecibosProvi!ReciboDefinitivo)
                        
                            If Val(WTiporeg) = 2 Then
                                If Val(WTipo2) = 2 And WEstado2 <> "X" And Val(WDefinitivo) = 0 Then
                            
                                    ZLugarII = ZLugarII + 1
                                    Auxi$ = Str$(rstRecibosProvi!Importe2)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibosProvi!Numero2))
                                    WFecha2 = IIf(IsNull(rstRecibosProvi!Fecha2), "", rstRecibosProvi!Fecha2)
                                    Call Ceros(Numero, 6)
                                    IngresaItem = Numero + "  " + rstRecibosProvi!Fecha2 + "  " + Auxi$ + "  " + rstRecibosProvi!Banco2
                                    
                                    WOrdFecha2 = IIf(IsNull(rstRecibosProvi!FechaOrd2), "", rstRecibosProvi!FechaOrd2)
                                
                                    ZEntraII(ZLugarII, 1) = IngresaItem
                                    ZEntraII(ZLugarII, 2) = "2" + rstRecibosProvi!Clave
                                    ZEntraII(ZLugarII, 3) = WOrdFecha2
                                    ZEntraII(ZLugarII, 4) = rstRecibosProvi!Numero2
                                    ZEntraII(ZLugarII, 5) = WFecha2
                                
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibosProvi.Close
            
            End If
            
            ZZTotal = ZLugarI + ZLugarII
            ZLugarI = 0
            ZLugarII = 0
            
            For ZCicla = 1 To ZZTotal
            
                If ZEntraI(ZLugarI + 1, 1) <> "" And ZEntraII(ZLugarII + 1, 1) <> "" Then
                
                    If ZEntraI(ZLugarI + 1, 3) < ZEntraII(ZLugarII + 1, 3) Then
                
                        ZLugarI = ZLugarI + 1
                        IngresaItem = ZEntraI(ZLugarI, 1)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = ZEntraI(ZLugarI, 2)
                        WIndice.AddItem IngresaItem
                        
                            Else
                
                        ZLugarII = ZLugarII + 1
                        
                        ZZNumero2 = ZEntraII(ZLugarII, 4)
                        ZZFecha2 = ZEntraII(ZLugarII, 5)
                        
                        ZSql = ""
                        ZSql = ZSql + "Select Recibos.Numero2, Recibos.Fecha2"
                        ZSql = ZSql + " FROM Recibos"
                        ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero2 + "'"
                        ZSql = ZSql + " and Recibos.Fecha2 = " + "'" + ZZFecha2 + "'"
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        If rstRecibos.RecordCount > 0 Then
                            rstRecibos.Close
                                Else
                            IngresaItem = ZEntraII(ZLugarII, 1)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = ZEntraII(ZLugarII, 2)
                            WIndice.AddItem IngresaItem
                        End If
                        
                    End If
                    
                        Else
                
                    If ZEntraI(ZLugarI + 1, 1) <> "" Then
                        ZLugarI = ZLugarI + 1
                        IngresaItem = ZEntraI(ZLugarI, 1)
                        Pantalla.AddItem IngresaItem
                        IngresaItem = ZEntraI(ZLugarI, 2)
                        WIndice.AddItem IngresaItem
                    End If
                
                    If ZEntraII(ZLugarII + 1, 1) <> "" Then
                        ZLugarII = ZLugarII + 1
                        
                        ZZNumero2 = ZEntraII(ZLugarII, 4)
                        ZZFecha2 = ZEntraII(ZLugarII, 5)
                        
                        ZSql = ""
                        ZSql = ZSql + "Select Recibos.Numero2, Recibos.Fecha2"
                         ZSql = ZSql + " FROM Recibos"
                        ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero2 + "'"
                        ZSql = ZSql + " and Recibos.Fecha2 = " + "'" + ZZFecha2 + "'"
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        If rstRecibos.RecordCount > 0 Then
                            rstRecibos.Close
                                Else
                            IngresaItem = ZEntraII(ZLugarII, 1)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = ZEntraII(ZLugarII, 2)
                            WIndice.AddItem IngresaItem
                        End If
                        
                    End If
                    
                End If
                
            Next ZCicla
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Tipo = '50'"
            ZSql = ZSql + " and CtaCte.Saldo <> 0"
            ZSql = ZSql + " Order by CtaCte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
            Rem spCtacte = "ListaCtacte"
            Rem Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstCtacte.RecordCount Then
                With rstCtaCte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Val(rstCtaCte!Tipo) = 50 Then
                                WSaldo = rstCtaCte!Saldo
                                Call Redondeo(WSaldo)
                                If WSaldo <> 0 And rstCtaCte!Cliente <> Space$(6) Then
                                    Auxi$ = Str$(Abs(rstCtaCte!Saldo))
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    WNumero = rstCtaCte!Numero
                                    Call Ceros(WNumero, 6)
                                    IngresaItem = WNumero + "  " + rstCtaCte!Vencimiento1 + "  " + Auxi$ + "  " + rstCtaCte!Cliente
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCtaCte!Clave
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtaCte.Close
                
            End If
            
        Case 4
            spCuenta = "ListaCuentas"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount Then
            
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCuenta!Cuenta + "  " + rstCuenta!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCuenta!Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
                
            End If
                
     
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            Proveedor.Text = Claveven$
            spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                DesProveedor.Caption = RstProveedor!Nombre
                WPrvDireccion = RstProveedor!Direccion
                WPrvCuit = RstProveedor!Cuit
                WPrvIb = RstProveedor!NroIb
                WTipoprv = Val(RstProveedor!Tipo) + 1
                WTipoiva = Val(RstProveedor!Iva)
                WTipoIb = RstProveedor!CodIb
                RstProveedor.Close
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
            Proveedor.SetFocus
            
        Case 1
            If Tipo1.Value = True Then
            
                Entra = "S"
                Indice = Pantalla.ListIndex
                Compara1 = WIndice.List(Indice)
        
                For iRow = 1 To 12
                    Compara2 = Proveedor.Text + WVector1.TextMatrix(iRow, 2)
                    Compara2 = Compara2 + WVector1.TextMatrix(iRow, 1)
                    Compara2 = Compara2 + WVector1.TextMatrix(iRow, 3)
                    Compara2 = Compara2 + WVector1.TextMatrix(iRow, 4)
                    If Compara1 = Compara2 Then
                        Entra = "N"
                        Exit For
                    End If
                Next iRow
            
                If Entra = "S" Then
            
                    Entra = "N"
                    For iRow = 1 To 12
                        If WVector1.TextMatrix(iRow, 1) = "" Then
                            XRow = iRow
                            Entra = "S"
                            Exit For
                        End If
                    Next iRow
                    
                    If Entra = "N" Then
                        m$ = "La cantidad de facturas a cancelar supera las 12"
                        A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                        Exit Sub
                    End If
            
                    Indice = Pantalla.ListIndex
                    ClaveCtaprv = WIndice.List(Indice)
                    spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                    Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                    If RstCtaPrv.RecordCount > 0 Then
                        XTipo = RstCtaPrv!Tipo
                        XLetra = RstCtaPrv!Letra
                        XPunto = RstCtaPrv!Punto
                        XNumero = RstCtaPrv!Numero
                        XSaldo = RstCtaPrv!Saldo
                        XNroInterno = Str$(RstCtaPrv!NroInterno)
                        RstCtaPrv.Close
                    End If
            
                    WVector1.Row = XRow
                    
                    WVector1.Col = 1
                    WVector1.Text = XTipo
                    
                    WVector1.Col = 2
                    WVector1.Text = XLetra
                    
                    WVector1.Col = 3
                    WVector1.Text = XPunto
                    
                    WVector1.Col = 4
                    WVector1.Text = XNumero
                    
                    WVector1.Col = 5
                    WVector1.Text = XSaldo
                    WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                    WVector1.Col = 6
                    Select Case Val(XTipo)
                        Case 1
                            WVector1.Text = "Pago Factura nro. " + Str$(XNumero)
                        Case 2
                            WVector1.Text = "Pago Nota de Debito nro. " + Str$(XNumero)
                        Case 3
                            WVector1.Text = "Pago Nota de Credito nro. " + Str$(XNumero)
                        Case 5
                            WVector1.Text = "Anticipo nro. " + Str$(XNumero)
                        Case Else
                            WVector1.Text = ""
                    End Select
                    
                    WVector1.Col = 1
                    WVector1.Text = XTipo
            
                    spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                        XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                        rstIvaComp.Close
                    End If
            
                    ParidadTotal = 0
                    If XPAgo = 2 Then
            
                     If WEmpresa = "0001" Then
            
                            spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
                            Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                                 If rstCambio.RecordCount > 0 Then
                                     ParidadTotal = rstCambio!Cambio
                                     rstCambio.Close
                                 End If
                          Else
                     Rem para pellital 10-7-2013
                             spCambio = "ConsultaCambioadm " + "'" + FechaParidad.Text + "'"
                             Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                                   If rstCambio.RecordCount > 0 Then
                                       ParidadTotal = rstCambio!Cambio
                                       rstCambio.Close
                                   End If
                                          
                     End If
                        
                        
                        WSaldo = XSaldo
                        WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                        WDife = WSaldoUs - WSaldo
                        Call Redondeo(WDife)
                
                        If WDife <> 0 Then
                            If WDife > 0 Then
                
                                XRow = XRow + 1
                                WVector1.Row = XRow
                                
                                WVector1.Col = 1
                                WVector1.Text = "02"
                                WVector1.Col = 2
                                WVector1.Text = XLetra
                                WVector1.Col = 3
                                WVector1.Text = XPunto
                                WVector1.Col = 4
                                WVector1.Text = "99999999"
                                WVector1.Col = 5
                                WVector1.Text = Str$(WDife)
                                WVector1.Col = 6
                                WVector1.Text = "N/D por Diferencia de Cambio "
                                WVector1.Col = 1
                                WVector1.Text = "02"
                        
                                    Else
                
                                XRow = XRow + 1
                                WVector1.Row = XRow
                                
                                WVector1.Col = 1
                                WVector1.Text = "03"
                                WVector1.Col = 2
                                WVector1.Text = XLetra
                                WVector1.Col = 3
                                WVector1.Text = XPunto
                                WVector1.Col = 4
                                WVector1.Text = "99999999"
                                WVector1.Col = 5
                                WVector1.Text = Str$(WDife)
                                WVector1.Col = 6
                                WVector1.Text = "N/C por Diferencia de Cambio "
                                WVector1.Col = 1
                                WVector1.Text = "03"
                        
                            End If
                        End If
                
                
             
               
                   End If
               
                End If
            
            End If
            
            Call Suma_Datos
            
        Case 2
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 1 To 12
                Compara2 = WVector1.TextMatrix(iRow, 13)
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
                Entra = "N"
                For iRow = 1 To 12
                    If WVector1.TextMatrix(iRow, 7) = "" Then
                        XRow = iRow
                        Entra = "S"
                        Exit For
                    End If
                Next iRow
                
                If Entra = "N" Then
                    m$ = "La cantidad de valores entregados supera los 12"
                    A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                    Exit Sub
                End If
        
                Indice = Pantalla.ListIndex
                TipoRecibos = Left$(WIndice.List(Indice), 1)
                ClaveRecibos = Mid$(WIndice.List(Indice), 2, 10)
                
                If TipoRecibos = "1" Then
                
                    spRecibos = "ConsultaRecibosClave " + "'" + ClaveRecibos + "'"
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibos.RecordCount > 0 Then
                
                        WVector1.Row = XRow
                    
                        WVector1.Col = 7
                        If XIndice = 2 Then
                            WVector1.Text = "3"
                                Else
                            WVector1.Text = "4"
                        End If
                    
                        WVector1.Col = 8
                        WVector1.Text = rstRecibos!Numero2
                    
                        WVector1.Col = 9
                        WVector1.Text = rstRecibos!Fecha2
                    
                        WVector1.Col = 10
                        WVector1.Text = ""
                    
                        WVector1.Col = 11
                        WVector1.Text = rstRecibos!Banco2
                    
                        WVector1.Col = 12
                        WVector1.Text = rstRecibos!Importe2
                        WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                        WVector1.Col = 13
                        WVector1.Text = "1" + ClaveRecibos
                    
                        WVector1.Col = 14
                        WVector1.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
                    
                        WVector1.Col = 7
                        If XIndice = 2 Then
                            WVector1.Text = "3"
                                Else
                            WVector1.Text = "4"
                        End If
                        
                        rstRecibos.Close
                    
                        Call Suma_Datos
                    
                        Pantalla.List(Indice) = ""
                    
                    End If
                
                    If XRow < 12 Then
                        XRow = XRow + 1
                    End If
                    
                        Else
                        
                    Sql1 = "Select *"
                    Sql2 = " FROM RecibosProvi"
                    Sql3 = " Where RecibosProvi.Clave = " + "'" + ClaveRecibos + "'"
                    spRecibosProvi = Sql1 + Sql2 + Sql3
                    Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibosProvi.RecordCount > 0 Then
                
                        WVector1.Row = XRow
                    
                        WVector1.Col = 7
                        If XIndice = 2 Then
                            WVector1.Text = "3"
                                Else
                            WVector1.Text = "4"
                        End If
                    
                        WVector1.Col = 8
                        WVector1.Text = rstRecibosProvi!Numero2
                    
                        WVector1.Col = 9
                        WVector1.Text = rstRecibosProvi!Fecha2
                    
                        WVector1.Col = 10
                        WVector1.Text = ""
                    
                        WVector1.Col = 11
                        WVector1.Text = rstRecibosProvi!Banco2
                    
                        WVector1.Col = 12
                        WVector1.Text = rstRecibosProvi!Importe2
                        WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                        WVector1.Col = 13
                        WVector1.Text = "2" + ClaveRecibos
                    
                        WVector1.Col = 14
                        WVector1.Text = IIf(IsNull(rstRecibosProvi!Cuit), "", rstRecibosProvi!Cuit)
                    
                        WVector1.Col = 7
                        If XIndice = 2 Then
                            WVector1.Text = "3"
                                Else
                            WVector1.Text = "4"
                        End If
                        
                        rstRecibosProvi.Close
                    
                        Call Suma_Datos
                    
                        Pantalla.List(Indice) = ""
                    
                    End If
                
                    If XRow < 12 Then
                        XRow = XRow + 1
                    End If
                
                End If
            
            End If
            
            Rem WVector1.Col = 7
            Rem Call StartEdit
            
        Case 3
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For iRow = 1 To 12
                Compara2 = WVector1.TextMatrix(iRow, 13)
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next iRow
            
            If Entra = "S" Then
            
                Entra = "N"
                For iRow = 1 To 12
                    If WVector1.TextMatrix(iRow, 7) = "" Then
                        XRow = iRow
                        Entra = "S"
                        Exit For
                    End If
                Next iRow
                
                If Entra = "N" Then
                    m$ = "La cantidad de valores entregados supera los 12"
                    A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                    Exit Sub
                End If
            
                Indice = Pantalla.ListIndex
                ClaveCtacte = WIndice.List(Indice)
                spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                
                    WVector1.Row = XRow
                    
                    WVector1.Col = 7
                    WVector1.Text = "4"
                    
                    WVector1.Col = 8
                    WVector1.Text = rstCtaCte!Numero
                    
                    WVector1.Col = 9
                    WVector1.Text = rstCtaCte!Vencimiento1
                    
                    WVector1.Col = 10
                    WVector1.Text = ""
                    
                    WVector1.Col = 11
                    WVector1.Text = ""
                    
                    WVector1.Col = 12
                    WVector1.Text = rstCtaCte!Saldo
                    WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                    WVector1.Col = 13
                    WVector1.Text = ClaveCtacte
                    
                    WVector1.Col = 7
                    WVector1.Text = "4"
                    
                    rstCtaCte.Close
                    
                    Call Suma_Datos
                    
                    Pantalla.List(Indice) = ""
                    
                End If
                
                If XRow < 12 Then
                    XRow = XRow + 1
                End If
            
            End If
            
            Rem WVector1.Col = 7
            Rem Call StartEdit
            
        Case 4
            Rem Indice = Pantalla.ListIndex
            Rem ClaveCuenta = WIndice.List(Indice)
            Rem spCuenta = "ConsultaCuentas " + "'" + ClaveCuenta + "'"
            Rem Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstCuenta.RecordCount > 0 Then
                
            Pantalla.Visible = False
            Rem Cuenta.SetFocus
                
        Case Else
    End Select
    
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
            spCuenta = "ConsultaCuentas " + "'" + Cuenta.Text + "'"
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                If WProceso = 0 Then
                    WCuenta(WVector1.Row, 1) = Cuenta.Text
                    IngreCuenta.Visible = False
                    WVector1.Row = WVector1.Row + 1
                    WVector1.Col = 1
                    Call StartEdit
                        Else
                    WCuenta(WVector1.Row, 2) = Cuenta.Text
                    IngreCuenta.Visible = False
                    WVector1.TextMatrix(WVector1.Row, 8) = ""
                    WVector1.TextMatrix(WVector1.Row, 9) = ""
                    WVector1.TextMatrix(WVector1.Row, 10) = ""
                    WVector1.TextMatrix(WVector1.Row, 11) = "Varios"
                    WVector1.Col = 12
                    Call StartEdit
                End If
                WProceso = 0
                    Else
                Cuenta.SetFocus
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Load()


    
    
    Call Limpia_Vector
    
    TipoPago.Clear
    
    TipoPago.AddItem "Normal"
    TipoPago.AddItem "Cheque Rechazado"
    
    TipoPago.ListIndex = 0
     
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Tipo6.Value = False
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FechaParidad.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    Tipo6.Value = False
    Debitos.Caption = ""
    Creditos.Caption = ""
    Dife.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    RetIb.Text = ""
    RetIbCiudad.Text = ""
    RetIva.Text = ""
    
    Carpeta.Text = ""
    Carpeta1.Text = ""
    Carpeta2.Text = ""
    Carpeta3.Text = ""
    Carpeta4.Text = ""
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    WLeyenda(6) = "Locacion de Obras y/o servicios"
    WLeyenda(7) = "Transporte de Carga"
    WLeyenda(8) = "Factura M"
    
    WParametro(0) = 0
    WParametro(1) = 2000
    WParametro(2) = 4000
    WParametro(3) = 8000
    WParametro(4) = 14000
    WParametro(5) = 24000
    WParametro(6) = 1000000
    
    WTasa1(1) = 0.1
    WTasa1(2) = 0.14
    WTasa1(3) = 0.18
    WTasa1(4) = 0.22
    WTasa1(5) = 0.26
    WTasa1(6) = 0.26
    
   If WEmpresa = "0008" Then
       spCambio = "ConsultaCambioadm " + "'" + FechaParidad.Text + "'"
       Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
       If rstCambio.RecordCount > 0 Then
           Paridad.Text = Str$(rstCambio!Cambio)
           Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
           rstCambio.Close
        End If
  
      Else
        spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
        Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambio.RecordCount > 0 Then
               Paridad.Text = Str$(rstCambio!Cambio)
               Paridad.Text = Pusing("#,###,###.####", Paridad.Text)
                 rstCambio.Close
             End If
 End If
    
    
    Orden.Text = ""
    Rem spPagos = "ListaPagosNumero"
    Rem Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPagos.RecordCount > 0 Then
    Rem     With rstPagos
    Rem         .MoveLast
    Rem         Orden.Text = rstPagos!Orden + 1
    Rem     End With
    Rem     rstPagos.Close
    Rem End If
    
End Sub

Private Sub IMPREORDEN()

    On Error GoTo WError
    
    ZSql = "DELETE PagosResumen"
    spPagosResumen = ZSql
    Set rstPagosResumen = db.OpenRecordset(spPagosResumen, dbOpenSnapshot, dbSQLPassThrough)
    LugarResumen = 0
        
    DA = 0
    With rstImprePago
        .Index = "Orden"
        .Seek ">=", 0
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

    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = WEmpresa
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Impretit = !Nombre
                Else
            Impretit = ""
        End If
    End With
    
    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpCuit = "30-54916508-3"
        Case Else
            WEmpCuit = "30-61052459-8"
    End Select
    
    
    ZZRazon = ""
    ZZCuitProveedor = ""
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        ZZRazon = RstProveedor!Nombre
        ZZCuitProveedor = RstProveedor!Cuit
        RstProveedor.Close
    End If
    
        
    WRenglon = 0
    Cantidad = 0
    Total = 0
    SubTotaL = 0
        
    Erase WImpresion, WDebito, WCredito, WImpre2
        
    For iRow = 1 To 12
    
        WRow = iRow
        
        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
        
            Cantidad = Cantidad + 1
            
            Select Case Val(Left$(WVector1.TextMatrix(iRow, 1), 2))
                Case 1
                    WImpresion(Cantidad, 2) = "Factura"
                Case 2
                    WImpresion(Cantidad, 2) = "N.Debito"
                Case 3
                    WImpresion(Cantidad, 2) = "N.Credito"
                Case 99
                    WImpresion(Cantidad, 2) = "Varios"
                Case Else
                    WImpresion(Cantidad, 2) = ""
            End Select
                            
            WImpresion(Cantidad, 3) = Left$(WVector1.TextMatrix(iRow, 4), 8)
            WImpresion(Cantidad, 4) = WVector1.TextMatrix(iRow, 6)
            WImpresion(Cantidad, 5) = WVector1.TextMatrix(iRow, 5)
            If Val(WImpresion(Cantidad, 2)) = 3 Or Val(WImpresion(Cantidad, 2)) = 5 Then
                Total = Total - Val(WImpresion(Cantidad, 5))
                    Else
                Total = Total + Val(WImpresion(Cantidad, 5))
            End If
                    
            WTipo = WVector1.TextMatrix(iRow, 1)
            WLetra = WVector1.TextMatrix(iRow, 2)
            WPunto = WVector1.TextMatrix(iRow, 3)
            WNumero = WVector1.TextMatrix(iRow, 4)
                
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                WImpresion(Cantidad, 1) = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                WImpresion(Cantidad, 1) = ""
            End If
                    
        End If
    Next iRow
        
    With rstEmpresa
        .Index = "Empresa"
        Claveven$ = "1"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            WCtaProveedor = !CtaProveedores
            WCtaEfectivo = !CtaEfectivo
            WCtaCheques = !CtaCheque
        End If
    End With
        
    If Tipo1.Value = True Or Tipo2.Value = True Then
        WDebito(1, 1) = WCtaProveedor
        WDebito(1, 2) = Total
            Else
        For iRow = 1 To 12
            WRow = iRow
            If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
                WDebito(iRow + 1, 1) = WCuenta(iRow, 1)
                WDebito(iRow + 1, 2) = Val(WVector1.TextMatrix(iRow, 5))
            End If
        Next iRow
    End If

    WCredito(1, 1) = WCtaProveedor
    If Retenido <> 0 Then
        WCredito(1, 2) = Retenido
    End If
        
    Lugar = 1
    Impre2 = 0
    SumaTercero = 0
    
    Rem If Val(Proveedor.Text) <> 0 Then
        For iRow = 1 To 12
            If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Then
                If Val(WVector1.TextMatrix(iRow, 7)) = 3 Then
                
                    SumaTercero = SumaTercero + Val(WVector1.TextMatrix(iRow, 12))
                    
                    LugarResumen = LugarResumen + 1
                    
                    ZZCorte = "1"
                    ZZOrden = Orden.Text
                    ZZRenglon = Str$(LugarResumen)
                    ZZProveedor = Proveedor.Text
                    ZZFecha = Fecha.Text
                    ZZImporteCheque = WVector1.TextMatrix(iRow, 12)
                    ZZNumeroCheque = Left$(WVector1.TextMatrix(iRow, 8), 8)
                    ZZFechaCheque = Left$(WVector1.TextMatrix(iRow, 9), 10)
                    ZZBancoCheque = Left$(WVector1.TextMatrix(iRow, 11), 20)
                    ZZCuit = WVector1.TextMatrix(iRow, 14)
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO PagosResumen ("
                    ZSql = ZSql + "Corte ,"
                    ZSql = ZSql + "Orden ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Proveedor ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "ImporteCheque ,"
                    ZSql = ZSql + "NumeroCheque ,"
                    ZSql = ZSql + "FechaCheque ,"
                    ZSql = ZSql + "BancoCheque ,"
                    ZSql = ZSql + "Razon ,"
                    ZSql = ZSql + "CuitProveedor )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZCorte + "',"
                    ZSql = ZSql + "'" + ZZOrden + "',"
                    ZSql = ZSql + "'" + ZZRenglon + "',"
                    ZSql = ZSql + "'" + ZZProveedor + "',"
                    ZSql = ZSql + "'" + ZZFecha + "',"
                    ZSql = ZSql + "'" + ZZCuit + "',"
                    ZSql = ZSql + "'" + ZZImporteCheque + "',"
                    ZSql = ZSql + "'" + ZZNumeroCheque + "',"
                    ZSql = ZSql + "'" + ZZFechaCheque + "',"
                    ZSql = ZSql + "'" + ZZBancoCheque + "',"
                    ZSql = ZSql + "'" + ZZRazon + "',"
                    ZSql = ZSql + "'" + ZZCuitProveedor + "')"
                    
                    spPagosResumen = ZSql
                    Set rstPagosResumen = db.OpenRecordset(spPagosResumen, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZZCorte = "2"
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO PagosResumen ("
                    ZSql = ZSql + "Corte ,"
                    ZSql = ZSql + "Orden ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Proveedor ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "ImporteCheque ,"
                    ZSql = ZSql + "NumeroCheque ,"
                    ZSql = ZSql + "FechaCheque ,"
                    ZSql = ZSql + "BancoCheque ,"
                    ZSql = ZSql + "Razon ,"
                    ZSql = ZSql + "CuitProveedor )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZCorte + "',"
                    ZSql = ZSql + "'" + ZZOrden + "',"
                    ZSql = ZSql + "'" + ZZRenglon + "',"
                    ZSql = ZSql + "'" + ZZProveedor + "',"
                    ZSql = ZSql + "'" + ZZFecha + "',"
                    ZSql = ZSql + "'" + ZZCuit + "',"
                    ZSql = ZSql + "'" + ZZImporteCheque + "',"
                    ZSql = ZSql + "'" + ZZNumeroCheque + "',"
                    ZSql = ZSql + "'" + ZZFechaCheque + "',"
                    ZSql = ZSql + "'" + ZZBancoCheque + "',"
                    ZSql = ZSql + "'" + ZZRazon + "',"
                    ZSql = ZSql + "'" + ZZCuitProveedor + "')"
                    
                    spPagosResumen = ZSql
                    Set rstPagosResumen = db.OpenRecordset(spPagosResumen, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            End If
        Next iRow
    Rem End If
    
    For iRow = 1 To 12
        If Val(WVector1.TextMatrix(iRow, 12)) <> 0 Then
            Lugar = Lugar + 1
            WCredito(Lugar, 4) = WVector1.TextMatrix(iRow, 12)
            Select Case Val(WVector1.TextMatrix(iRow, 7))
                Case 2
                    WCredito(Lugar, 1) = "999999"
                    ClaveBanco = WVector1.TextMatrix(iRow, 10)
                    spBanco = "ConsultaBanco " + "'" + ClaveBanco + "'"
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                        WCredito(Lugar, 1) = rstBanco!Cuenta
                        rstBanco.Close
                    End If
                Case 3, 4
                    WCredito(Lugar, 1) = WCtaCheques
                Case Else
                    WCredito(Lugar, 1) = WCtaEfectivo
            End Select
                    
            If Val(WVector1.TextMatrix(iRow, 7)) <> 3 Or Val(Proveedor.Text) = 0 Or Val(WEmpresa) = 8 Then
                Impre2 = Impre2 + 1
            
                WImpre2(Impre2, 1) = WVector1.TextMatrix(iRow, 8)
                WImpre2(Impre2, 2) = WVector1.TextMatrix(iRow, 11)
                WImpre2(Impre2, 3) = WVector1.TextMatrix(iRow, 12)
                WImpre2(Impre2, 4) = WVector1.TextMatrix(iRow, 9)
                    
                WCredito(Lugar, 2) = WVector1.TextMatrix(iRow, 11)
                WCredito(Lugar, 3) = WVector1.TextMatrix(iRow, 8)
                WCredito(Lugar, 4) = WVector1.TextMatrix(iRow, 12)
            End If
        End If
    Next iRow
    
    If SumaTercero <> 0 Then
        For WCiclo = 1 To 12
            If Val(WImpre2(WCiclo, 3)) = 0 Then
                If Val(WEmpresa) = 1 Then
                    WImpre2(WCiclo, 1) = ""
                    WImpre2(WCiclo, 2) = "Valores S/Detalle"
                    WImpre2(WCiclo, 3) = Str$(SumaTercero)
                    WImpre2(WCiclo, 4) = ""
                    Exit For
                End If
            End If
        Next WCiclo
    End If
    
    If LugarResumen > 0 Then
    
        ZDesde = LugarResumen
        
        For ZCiclo = ZDesde To 17
        
            LugarResumen = LugarResumen + 1
                    
            ZZCorte = "1"
            ZZOrden = Orden.Text
            ZZRenglon = Str$(LugarResumen)
            ZZProveedor = Proveedor.Text
            ZZFecha = Fecha.Text
            ZZImporteCheque = ""
            ZZNumeroCheque = ""
            ZZFechaCheque = ""
            ZZBancoCheque = ""
            ZZCuit = ""
                    
            ZSql = ""
            ZSql = ZSql + "INSERT INTO PagosResumen ("
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "ImporteCheque ,"
            ZSql = ZSql + "NumeroCheque ,"
            ZSql = ZSql + "FechaCheque ,"
            ZSql = ZSql + "BancoCheque ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "CuitProveedor )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZImporteCheque + "',"
            ZSql = ZSql + "'" + ZZNumeroCheque + "',"
            ZSql = ZSql + "'" + ZZFechaCheque + "',"
            ZSql = ZSql + "'" + ZZBancoCheque + "',"
            ZSql = ZSql + "'" + ZZRazon + "',"
            ZSql = ZSql + "'" + ZZCuitProveedor + "')"
                    
            spPagosResumen = ZSql
            Set rstPagosResumen = db.OpenRecordset(spPagosResumen, dbOpenSnapshot, dbSQLPassThrough)
                    
            ZZCorte = "2"
                    
            ZSql = ""
            ZSql = ZSql + "INSERT INTO PagosResumen ("
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "ImporteCheque ,"
            ZSql = ZSql + "NumeroCheque ,"
            ZSql = ZSql + "FechaCheque ,"
            ZSql = ZSql + "BancoCheque ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "CuitProveedor )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZFecha + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZImporteCheque + "',"
            ZSql = ZSql + "'" + ZZNumeroCheque + "',"
            ZSql = ZSql + "'" + ZZFechaCheque + "',"
            ZSql = ZSql + "'" + ZZBancoCheque + "',"
            ZSql = ZSql + "'" + ZZRazon + "',"
            ZSql = ZSql + "'" + ZZCuitProveedor + "')"
                    
            spPagosResumen = ZSql
            Set rstPagosResumen = db.OpenRecordset(spPagosResumen, dbOpenSnapshot, dbSQLPassThrough)
        
        Next ZCiclo
    End If
        
    SubTotaL = Total - Retenido
    TotalDebito = Total
    TotalCredito = Total

    For WCiclo = 1 To 12
    
        WFecha1 = ""
        WNumero1 = ""
        WComprobante1 = ""
        WDescripcion1 = ""
        WImporte1 = 0
        WNumero2 = ""
        WBanco2 = ""
        WImporte2 = 0
        WFecha2 = ""
            
        If Val(WImpresion(WCiclo, 5)) <> 0 Then
            WFecha1 = WImpresion(WCiclo, 1)
            WNumero1 = WImpresion(WCiclo, 3)
            WComprobante1 = WImpresion(WCiclo, 2)
            WDescripcion1 = WImpresion(WCiclo, 4)
            WImporte1 = Val(WImpresion(WCiclo, 5))
        End If
                    
        If Val(WImpre2(WCiclo, 3)) <> 0 Then
            WNumero2 = WImpre2(WCiclo, 1)
            WBanco2 = WImpre2(WCiclo, 2)
            WImporte2 = Val(WImpre2(WCiclo, 3))
            WFecha2 = WImpre2(WCiclo, 4)
        End If
        
        WRenglon = WRenglon + 1
        With rstImprePago
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = WRenglon
            !Fecha = Fecha.Text
            !Proveedor = Proveedor.Text
            !Nombre = DesProveedor.Caption
            !Fecha1 = WFecha1
            !Numero1 = WNumero1
            !Comprobante1 = WComprobante1
            !Descripcion1 = WDescripcion1
            !Importe1 = WImporte1
            !Numero2 = WNumero2
            !Banco2 = WBanco2
            !Importe2 = WImporte2
            !Fecha2 = WFecha2
            !Neto = Total
            !Rete1 = Val(Retencion.Text)
            !Rete2 = Val(RetIb.Text) + Val(RetIbCiudad.Text)
            !Total = Val(RetIva.Text)
            !Observaciones = Observaciones.Text
            !Empresa = Impretit
            !Cuit = WEmpCuit
            !Paridad = ParidadTotal
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = WRenglon
            !Fecha = Fecha.Text
            !Proveedor = Proveedor.Text
            !Nombre = DesProveedor.Caption
            !Fecha1 = WFecha1
            !Numero1 = WNumero1
            !Comprobante1 = WComprobante1
            !Descripcion1 = WDescripcion1
            !Importe1 = WImporte1
            !Numero2 = WNumero2
            !Banco2 = WBanco2
            !Importe2 = WImporte2
            !Fecha2 = WFecha2
            !Neto = Total
            !Rete1 = Val(Retencion.Text)
            !Rete2 = Val(RetIb.Text) + Val(RetIbCiudad.Text)
            !Total = Val(RetIva.Text)
            !Observaciones = Observaciones.Text
            !Empresa = Impretit
            !Cuit = WEmpCuit
            !Paridad = ParidadTotal
            .Update
        End With
        
    Next WCiclo

    LISTADO.ReportFileName = "Imprepago.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
    
    
    Emite = "N"
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and Pagos.Tipo2 = " + "'" + "3" + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        Emite = "S"
        rstPagos.Close
    End If
    
    Rem If Emite = "S" And Val(Proveedor.Text) <> 0 And Val(WEmpresa) = 1 Then
    If Emite = "S" And Val(WEmpresa) = 1 Then
    
        Uno = "{PagosResumen.Orden} in " + Orden.Text + " to " + Orden.Text
        
        LISTADO.GroupSelectionFormula = Uno
        LISTADO.SelectionFormula = Uno
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        LISTADO.SQLQuery = "SELECT PagosResumen.Corte, PagosResumen.Orden, PagosResumen.Renglon, PagosResumen.Fecha, PagosResumen.Cuit, PagosResumen.ImporteCheque, PagosResumen.NumeroCheque, PagosResumen.FechaCheque, PagosResumen.BancoCheque, PagosResumen.Razon, PagosResumen.CuitProveedor " _
                    + "From " _
                    + DSQ + ".dbo.PagosResumen PagosResumen " _
                    + "Where " _
                    + "PagosResumen.Orden >= " + Orden.Text + " AND " _
                    + "PagosResumen.Orden <= " + Orden.Text
        
        LISTADO.Connect = Connect()
    
        LISTADO.ReportFileName = "ImpreAnexoOrdenII.rpt"
        LISTADO.Destination = 1
        LISTADO.CopiesToPrinter = 1
        LISTADO.Action = 1
        
        LISTADO.GroupSelectionFormula = ""
        LISTADO.SelectionFormula = ""
        
    End If
    
    Exit Sub
        
WError:
    Resume Next
  

End Sub


Private Sub Impreret()

    On Error GoTo WError
        
    WRenglon = 0
    DA = 0
    With rstImpreRetGan
        .Index = "Orden"
        .Seek ">=", 0
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
    
    Mes% = Val(Mid$(Fecha.Text, 3, 2))
    WCuatri = ""

    If Mes% <= 4 Then
        WCuatri = "Primer Cuatrimestre"
            Else
        If Mes% >= 5 And Mes% <= 8 Then
            WCuatri = "Segundo Cuatrimestre"
                Else
            If Mes% >= 9 Then
                WCuatri = "Tercer Cuatrimestre"
            End If
        End If
    End If

    Select Case Val(WEmpresa)
        Case 1
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
    End Select
    
    
    With rstImpreRetGan
        .AddNew
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        !Clave = "1" + Auxi + Auxi1
        !Tipo = 1
        !Orden = Val(Orden.Text)
        !Renglon = WRenglon
        !NroCertificado = WCertificadoGan
        !Empresa = WEmpNombre
        !Direccion = WEmpDireccion
        !Localidad = WEmpLocalidad
        !Fecha = Fecha.Text
        !Cuit = WEmpCuit
        !NombrePrv = DesProveedor.Caption
        !DireccionPrv = WPrvDireccion
        !CuitPrv = WPrvCuit
        !Concepto = WLeyenda$(Val(WTipoprv))
        !Pagado = Total - WRetencion
        !Retenido = WRetencion
        .Update
    End With
    
    With rstImpreRetGan
        .AddNew
        Auxi = Orden.Text
        Call Ceros(Auxi, 6)
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        !Clave = "2" + Auxi + Auxi1
        !Tipo = 2
        !Orden = Val(Orden.Text)
        !Renglon = WRenglon
        !NroCertificado = WCertificadoGan
        !Empresa = WEmpNombre
        !Direccion = WEmpDireccion
        !Localidad = WEmpLocalidad
        !Fecha = Fecha.Text
        !Cuit = WEmpCuit
        !NombrePrv = DesProveedor.Caption
        !DireccionPrv = WPrvDireccion
        !CuitPrv = WPrvCuit
        !Concepto = WLeyenda$(Val(WTipoprv))
        !Pagado = Total - WRetencion
        !Retenido = WRetencion
        .Update
    End With
        
    LISTADO.ReportFileName = "Impreretgan.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next
    

End Sub


Private Sub calcret_Click()

    WRetencion = 0
    Retencion.Text = ""
    RetIva.Text = ""
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WTipoiva = Val(RstProveedor!Iva)
        RstProveedor.Close
    End If
    
    If TipoPago.ListIndex = 0 Then
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Or WTipoprv = 6 Or WTipoprv = 7 Then
        
            Rem XBruto = Val(Debitos.Caption)
            XBruto = ZZBase
            If WTipoiva = 2 Then
                XNeto = (XBruto / 1.21)
                    Else
                XNeto = XBruto
            End If
            XIva = XBruto - XNeto
            XTBase = XNeto
        
            WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
            
            ClaveRetencion = WFecha + Proveedor.Text
            spRetencion = "ConsultaRetencion " + "'" + ClaveRetencion + "'"
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            If rstRetencion.RecordCount > 0 Then
                WNeto = rstRetencion!Neto
                WAnticipo = rstRetencion!Anticipo
                WBruto = rstRetencion!Bruto
                WIva = rstRetencion!Iva
                WRetenido = rstRetencion!Retenido
                rstRetencion.Close
                    Else
                XFecha = WFecha
                XProveedor = Proveedor.Text
                XXNeto = ""
                XXAnticipo = ""
                XXBruto = ""
                XXIva = ""
                XXRetenido = ""
                XClave = XFecha + XProveedor
                
                XParam = "'" + XClave + "','" _
                        + XFecha + "','" + XProveedor + "','" _
                        + XXNeto + "','" _
                        + XXRetenido + "','" + XXAnticipo + "','" _
                        + XXBruto + "','" _
                        + XXAcumulado + "'"
                    
                spRstRetencion = "AltaRetencion " + XParam
                Set RstRstRetencion = db.OpenRecordset(spRstRetencion, dbOpenSnapshot, dbSQLPassThrough)
                    
                WNeto = 0
                WAnticipo = 0
                WBruto = 0
                WIva = 0
                WRetenido = 0
            End If
            
            Select Case WTipoprv
                Case 1
                    WMinimo = 12000
                Case 2
                    WMinimo = 1200
                Case 3
                    WMinimo = 1200
                Case 6
                    WMinimo = 5000
                Case 7
                    WMinimo = 6500
                Case Else
            End Select

            WAcupag = WNeto + XTBase
            WAuxi = WAcupag - WMinimo

            If WAuxi <= 0 Then
                WAuxi = 0
                WRetencion = 0
            End If

            WTasa = 0.02
            If WTipoprv = 1 Then
                    WTasa = 0.02
            End If
            If WTipoprv = 3 Then
                    WTasa = 0.06
            End If
            If WTipoprv = 7 Then
                    WTasa = 0.0025
            End If

            Select Case WTipoprv
                Case 2
                    WRetencion = 0
                    WTope = 0
                    WTope1 = 0
                    
                    For DA = 0 To 5
                        If WAuxi >= WParametro(DA) And WAuxi < WParametro(DA + 1) Then
                            WTope1 = WAuxi
                            WTope = WParametro(DA)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(DA + 1)
                            WRetencion = WRetencion + WSum
                        End If
                        If WAuxi >= WParametro(DA + 1) Then
                            WTope1 = WParametro(DA + 1)
                            WTope = WParametro(DA)
                            WSum = WTope1 - WTope
                            WSum = WSum * WTasa1(DA + 1)
                            WRetencion = WRetencion + WSum
                        End If
                    Next DA
                    
                Case Else
                    WRetencion = WAuxi * WTasa
                    
            End Select

            WRetencion = WRetencion - WRetenido

            If WRetencion < 20 Then
                WRetencion = 0
                        Else
                If WRetencion > XNeto Then
                        WRetencion = 0
                End If
            End If
                    
            Call Redondeo(WRetencion)
            Retencion.Text = WRetencion
            Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            
        End If
        
        Rem If WTipoprv = 8 Then
        
            WRete1 = 0
            WRete2 = 0
                
            For iRow = 1 To 12
            
                XTipo = Left$(WVector1.TextMatrix(iRow, 1), 2)
                XLetra = Left$(WVector1.TextMatrix(iRow, 2), 1)
                XPunto = Left$(WVector1.TextMatrix(iRow, 3), 4)
                XNumero = Left$(WVector1.TextMatrix(iRow, 4), 8)
                XImporte = WVector1.TextMatrix(iRow, 5)
                
                If Val(XImporte) <> 0 And XLetra = "M" Then
                
                    XBruto = Val(XImporte)
                    XNeto = (XBruto / 1.21)
                    XIva = XBruto - XNeto

                    If XNeto >= 1000 Then
            
                        WTasa = 0.03
                        WRete1 = WRete1 + (XNeto * WTasa)
                    
                        Sql1 = "Select *"
                        Sql2 = " FROM IvaComp"
                        Sql3 = " Where IvaComp.Proveedor = " + "'" + Proveedor.Text + "'"
                        Sql4 = " and IvaComp.Tipo = " + "'" + XTipo + "'"
                        Sql5 = " and IvaComp.Letra = " + "'" + XLetra + "'"
                        Sql6 = " and IvaComp.Punto = " + "'" + XPunto + "'"
                        Sql7 = " and IvaComp.Numero = " + "'" + XNumero + "'"
                        spIvaComp = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
                        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                        If rstIvaComp.RecordCount > 0 Then
                            WRete2 = WRete2 + rstIvaComp!Iva21
                            rstIvaComp.Close
                        End If
                
                    End If
                    
                End If
                
            Next iRow
            
            If Val(Retencion.Text) = 0 Then
                Call Redondeo(WRete1)
                WRetencion = WRete1
                Retencion.Text = Str$(WRete1)
                Retencion.Text = Pusing("#,###,###.##", Retencion.Text)
            End If
            
            Call Redondeo(WRete2)
            RetIva.Text = Str$(WRete2)
            RetIva.Text = Pusing("#,###,###.##", RetIva.Text)
            
        Rem End If
        
    End If
    
    End If

End Sub


Private Sub CalcRetIb()

    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoIb = RstProveedor!CodIb
        WTipoiva = Val(RstProveedor!Iva)
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WPorceIb = IIf(IsNull(RstProveedor!PorceIb), "0", RstProveedor!PorceIb)
        WPorceIbCaba = IIf(IsNull(RstProveedor!PorceIbCaba), "0", RstProveedor!PorceIbCaba)
        RstProveedor.Close
    End If

    WRetIb = 0
    RetIb.Text = ""
    
    WRetIbCiudad = 0
    RetIbCiudad.Text = ""
    
    If TipoPago.ListIndex = 0 Then
    
        If Tipo1.Value = True Or Tipo2.Value = True Then
        
            If WTipoIb = 0 Or WTipoIb = 1 Then
            
                Rem XBruto = Val(Debitos.Caption)
                XBruto = ZZBase
                If WTipoiva = 2 Then
                    XNeto = (XBruto / 1.21)
                        Else
                    XNeto = XBruto
                End If
                XIva = XBruto - XNeto
                XTBase = XNeto
                                
                Rem If XTBase >= 400 Then
                
                    WImpoRetenido = 0
        
                    For iRow = 1 To 12
            
                        WLetra = WVector1.TextMatrix(iRow, 2)
                        WTipo = WVector1.TextMatrix(iRow, 1)
                        WPunto = WVector1.TextMatrix(iRow, 3)
                        WNumero = WVector1.TextMatrix(iRow, 4)
                                    
                        ZRechazado = 0
                        ZNroInterno = "0"
                        
                        ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM CtaCtePrv"
                        ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClaveCtaCtePrv + "'"
                        spCtaCtePrv = ZSql
                        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCtePrv.RecordCount > 0 Then
                            ZNroInterno = Str$(rstCtaCtePrv!NroInterno)
                            rstCtaCtePrv.Close
                        End If
                        
                        spIvaComp = "Consultaivacomp " + "'" + ZNroInterno + "'"
                        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                        If rstIvaComp.RecordCount > 0 Then
                            ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
                            rstIvaComp.Close
                        End If
                    
                        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
                        If (Val(WVector1.TextMatrix(iRow, 1)) <> 0 And ZRechazado = 0) Or Tipo2.Value = True Then
                                
                            WImpre4 = Val(WVector1.TextMatrix(iRow, 5))
                            If WTipoiva = 2 Then
                                WImpre4 = WImpre4 / 1.21
                            End If
                            Call Redondeo(WImpre4)
                            XImpre4 = Str$(WImpre4)
                            
                            ZFechaCompa = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            If ZFechaCompa >= "20071201" Then
                            
                                Select Case WTipoIb
                                    Case 0, 1
                                        WRete = Val(XImpre4) * (WPorceIb / 100)
                                        Call Redondeo(WRete)
                                        WImpoRetenido = WImpoRetenido + WRete
                                    Case Else
                                End Select
                                
                                    Else
                    
                                Select Case WTipoIb
                                    Case 0
                                        WRete = Val(XImpre4) * (0.75 / 100)
                                        Call Redondeo(WRete)
                                        WImpoRetenido = WImpoRetenido + WRete
                            
                                    Case Else
                                        WRete = Val(XImpre4) * (1.75 / 100)
                                        Call Redondeo(WRete)
                                        WImpoRetenido = WImpoRetenido + WRete
                                End Select
                                
                            End If
                        
                        End If
                        End If
                    Next iRow
        
                    WRetIb = WImpoRetenido
                    
                Rem End If
                
                Call Redondeo(WRetIb)
                RetIb.Text = WRetIb
                RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
                
                Rem XBruto = Val(Debitos.Caption)
                Rem If WTipoiva = 2 Then
                Rem     XNeto = (XBruto / 1.21)
                Rem         Else
                Rem     XNeto = XBruto
                Rem End If
                Rem XIva = XBruto - XNeto
                Rem XTBase = XNeto
                Rem
                Rem If XTBase >= 400 Then
                Rem     Select Case WTipoIb
                Rem         Case 0
                Rem             WRetIb = XTBase * (0.75 / 100)
                Rem         Case 1
                Rem             WRetIb = XTBase * (1.75 / 100)
                Rem         Case Else
                Rem             WRetIb = 0
                Rem     End Select
                Rem End If
                Rem
                Rem Call Redondeo(WRetIb)
                Rem RetIb.Text = WRetIb
                Rem RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
                
            End If
            
            If WTipoIb = 3 Or WTipoIb = 4 Then
            
                If Val(WEmpresa) = 1 Then
                
                    Rem XBruto = Val(Debitos.Caption)
                    XBruto = ZZBase
                    If WTipoiva = 2 Then
                        XNeto = (XBruto / 1.21)
                            Else
                        XNeto = XBruto
                    End If
                    XIva = XBruto - XNeto
                    XTBase = XNeto
                                    
                    Rem If XTBase >= 500 Then
                    
                        WImpoRetenido = 0
            
                        For iRow = 1 To 12
            
                            WLetra = WVector1.TextMatrix(iRow, 2)
                            WTipo = WVector1.TextMatrix(iRow, 1)
                            WPunto = WVector1.TextMatrix(iRow, 3)
                            WNumero = WVector1.TextMatrix(iRow, 4)
                                        
                            ZRechazado = 0
                            ZNroInterno = "0"
                            
                            ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
                            
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM CtaCtePrv"
                            ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClaveCtaCtePrv + "'"
                            spCtaCtePrv = ZSql
                            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCtaCtePrv.RecordCount > 0 Then
                                ZNroInterno = Str$(rstCtaCtePrv!NroInterno)
                                rstCtaCtePrv.Close
                            End If
                            
                            spIvaComp = "Consultaivacomp " + "'" + ZNroInterno + "'"
                            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                            If rstIvaComp.RecordCount > 0 Then
                                ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
                                rstIvaComp.Close
                            End If
                        
                        
                            If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then

                                WImpre4 = Val(WVector1.TextMatrix(iRow, 5))
                                If WTipoiva = 2 Then
                                    WImpre4 = WImpre4 / 1.21
                                End If
                                Call Redondeo(WImpre4)
                                XImpre4 = Str$(WImpre4)
                                
                                If Val(XImpre4) < 300 Then
                                    XImpre4 = "0"
                                End If
                                
                                If WPorceIbCaba <> 0 Then
                                    WRete = Val(XImpre4) * (WPorceIbCaba / 100)
                                    Call Redondeo(WRete)
                                    WImpoRetenido = WImpoRetenido + WRete
                                        Else
                                    If WTipoIb = 3 Then
                                        WRete = Val(XImpre4) * (2.5 / 100)
                                            Else
                                        WRete = Val(XImpre4) * (4.5 / 100)
                                    End If
                                    Call Redondeo(WRete)
                                    WImpoRetenido = WImpoRetenido + WRete
                                End If
                            
                            End If
                        Next iRow
                    
            
                        WRetIbCiudad = WImpoRetenido
                        
                    Rem End If
                    
                    Call Redondeo(WRetIbCiudad)
                    RetIbCiudad.Text = WRetIbCiudad
                    RetIbCiudad.Text = Pusing("#,###,###.##", RetIbCiudad.Text)
                    
                    Rem XBruto = Val(Debitos.Caption)
                    Rem If WTipoiva = 2 Then
                    Rem     XNeto = (XBruto / 1.21)
                    Rem         Else
                    Rem     XNeto = XBruto
                    Rem End If
                    Rem XIva = XBruto - XNeto
                    Rem XTBase = XNeto
                    Rem
                    Rem If XTBase >= 400 Then
                    Rem     Select Case WTipoIb
                    Rem         Case 0
                    Rem             WRetIb = XTBase * (0.75 / 100)
                    Rem         Case 1
                    Rem             WRetIb = XTBase * (1.75 / 100)
                    Rem         Case Else
                    Rem             WRetIb = 0
                    Rem     End Select
                    Rem End If
                    Rem
                    Rem Call Redondeo(WRetIb)
                    Rem RetIb.Text = WRetIb
                    Rem RetIb.Text = Pusing("#,###,###.##", RetIb.Text)
                    
                End If
            
            End If
            
        End If
    
    End If

End Sub

Private Sub Impreretib()

    On Error GoTo WError
        
    WRenglon = 0
    DA = 0
    With rstImpreRetIb
        .Index = "Orden"
        .Seek ">=", 0
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
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoIb = RstProveedor!CodIb
        WTipoiva = Val(RstProveedor!Iva)
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WPorceIb = IIf(IsNull(RstProveedor!PorceIb), "0", RstProveedor!PorceIb)
        WPorceIbCaba = IIf(IsNull(RstProveedor!PorceIbCaba), "0", RstProveedor!PorceIbCaba)
        RstProveedor.Close
    End If

    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
            WNroIb = "902-913585-2"
            WNroAgente = ""
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
            WNroIb = ""
            WNroAgente = ""
    End Select
    
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
        
    WImpoRetenido = 0
        
    For iRow = 1 To 12
    
        WRow = iRow
        
        WLetra = WVector1.TextMatrix(iRow, 2)
        WTipo = WVector1.TextMatrix(iRow, 1)
        WPunto = WVector1.TextMatrix(iRow, 3)
        WNumero = WVector1.TextMatrix(iRow, 4)
                    
        ZRechazado = 0
        ZNroInterno = "0"
        
        ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCtePrv"
        ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClaveCtaCtePrv + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCtePrv.RecordCount > 0 Then
            ZNroInterno = Str$(rstCtaCtePrv!NroInterno)
            rstCtaCtePrv.Close
        End If
        
        spIvaComp = "Consultaivacomp " + "'" + ZNroInterno + "'"
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
            rstIvaComp.Close
        End If
                    
        
        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
        If (Val(WVector1.TextMatrix(iRow, 1)) <> 0 And ZRechazado = 0) Or Tipo2.Value = True Then
        
            Select Case Val(WVector1.TextMatrix(iRow, 1))
                Case 1
                    XImpre1 = "Factura"
                Case 2
                    XImpre1 = "N.Debito"
                Case 3
                    XImpre1 = "N.Credito"
                Case 5
                    XImpre1 = "Anticipo"
                Case 99
                    XImpre1 = "Varios"
                Case Else
                    XImpre1 = ""
            End Select
                                
            XImpre2 = Left$(WVector1.TextMatrix(iRow, 4), 8)
                
            Rem spIvacomp = "ConsultaIvacomp " + "'" + XImpre2 + "'"
            Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstIvacomp.RecordCount > 0 Then
            Rem     XImpre2 = rstIvacomp!Numero
            Rem     rstIvacomp.Close
            Rem End If
                    
            WTipo = WVector1.TextMatrix(iRow, 1)
            WLetra = WVector1.TextMatrix(iRow, 2)
            WPunto = WVector1.TextMatrix(iRow, 3)
            WNumero = WVector1.TextMatrix(iRow, 4)
                    
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XImpre3 = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                XImpre3 = ""
            End If
                        
            WImpre4 = Val(WVector1.TextMatrix(iRow, 5))
            Rem If Val(WTipo) = 3 Or Val(WTipo) = 5 Then
            Rem    WImpre4 = WImpre4 * -1
            Rem End If
            If WTipoiva = 2 Then
                WImpre4 = WImpre4 / 1.21
            End If
            Call Redondeo(WImpre4)
            XImpre4 = Str$(WImpre4)
            
            ZFechaCompa = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            If ZFechaCompa >= "20071201" Then
            
                Select Case WTipoIb
                    Case 0, 1
                        WRete = Val(XImpre4) * (WPorceIb / 100)
                        Call Redondeo(WRete)
                        WImpoRetenido = WImpoRetenido + WRete
                            
                        WRenglon = WRenglon + 1
                        With rstImpreRetIb
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "1" + Auxi + Auxi1
                            !Tipo = 1
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION " + Str$(WPorceIb) + "%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = WPorceIb
                            !Retencion1 = WRete
                            .Update
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "2" + Auxi + Auxi1
                            !Tipo = 2
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION " + Str$(WPorceIb) + "%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = WPorceIb
                            !Retencion1 = WRete
                            .Update
                        End With
                    Case Else
                End Select
                    
                    Else
                    
                Select Case WTipoIb
                    Case 0
                        WRete = Val(XImpre4) * (0.75 / 100)
                        Call Redondeo(WRete)
                        WImpoRetenido = WImpoRetenido + WRete
                            
                        WRenglon = WRenglon + 1
                        With rstImpreRetIb
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "1" + Auxi + Auxi1
                            !Tipo = 1
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION 0.75%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = 0.75
                            !Retencion1 = WRete
                            .Update
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "2" + Auxi + Auxi1
                            !Tipo = 2
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION 0.75%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = 0.75
                            !Retencion1 = WRete
                            .Update
                        End With
                            
                    Case Else
                        WRete = Val(XImpre4) * (1.75 / 100)
                        Call Redondeo(WRete)
                        WImpoRetenido = WImpoRetenido + WRete
                    
                        WRenglon = WRenglon + 1
                        With rstImpreRetIb
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "1" + Auxi + Auxi1
                            !Tipo = 1
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION 1.75%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = 1.75
                            !Retencion1 = WRete
                            .Update
                            .AddNew
                            Auxi = Orden.Text
                            Call Ceros(Auxi, 6)
                            Auxi1 = WRenglon
                            Call Ceros(Auxi1, 2)
                            !Clave = "2" + Auxi + Auxi1
                            !Tipo = 2
                            !Orden = Val(Orden.Text)
                            !Renglon = WRenglon
                            !Empresa = WEmpNombre
                            !Direccion = WEmpDireccion
                            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
                            !Localidad = WEmpLocalidad
                            !Fecha = Fecha.Text
                            !Cuit = WEmpCuit
                            !Copia = ImpreCopia(DA)
                            !NroIb = WNroIb
                            !NroAgente = WNroAgente
                            !NombrePrv = DesProveedor.Caption
                            !DireccionPrv = WPrvDireccion
                            !CuitPrv = WPrvCuit
                            !NroIbPrv = WPrvIb
                            !Tipo1 = XImpre1
                            !Numero1 = XImpre2
                            !Fecha1 = XImpre3
                            !Categoria1 = "SUJETO A RETENCION 1.75%"
                            !Importe1 = Val(XImpre4)
                            !Porce1 = 1.75
                            !Retencion1 = WRete
                            .Update
                        End With
                    
                End Select
            End If
        End If
        End If
    Next iRow
    
    For Ciclo = WRenglon + 1 To 12
        With rstImpreRetIb
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIb)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
        End With
    Next Ciclo
    
    LISTADO.GroupSelectionFormula = ""
    LISTADO.SelectionFormula = ""
        
    LISTADO.ReportFileName = "Impreretib.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub


Private Sub ImpreretibCiudad()

    On Error GoTo WError
        
    WRenglon = 0
    DA = 0
    With rstImpreRetIb
        .Index = "Orden"
        .Seek ">=", 0
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
    
    spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        WTipoIb = RstProveedor!CodIb
        WTipoiva = Val(RstProveedor!Iva)
        WTipoprv = Val(RstProveedor!Tipo) + 1
        WPorceIb = IIf(IsNull(RstProveedor!PorceIb), "0", RstProveedor!PorceIb)
        WPorceIbCaba = IIf(IsNull(RstProveedor!PorceIbCaba), "0", RstProveedor!PorceIbCaba)
        RstProveedor.Close
    End If

    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
            WNroIb = "902-913585-2"
            WNroAgente = ""
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
            WNroIb = ""
            WNroAgente = ""
    End Select
    
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
        
    WImpoRetenido = 0
        
    For iRow = 1 To 12
        WRow = iRow
        
        WLetra = WVector1.TextMatrix(iRow, 2)
        WTipo = WVector1.TextMatrix(iRow, 1)
        WPunto = WVector1.TextMatrix(iRow, 3)
        WNumero = WVector1.TextMatrix(iRow, 4)
                    
        ZRechazado = 0
        ZNroInterno = "0"
        
        ZZClaveCtaCtePrv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCtePrv"
        ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClaveCtaCtePrv + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCtePrv.RecordCount > 0 Then
            ZNroInterno = Str$(rstCtaCtePrv!NroInterno)
            rstCtaCtePrv.Close
        End If
        
        spIvaComp = "Consultaivacomp " + "'" + ZNroInterno + "'"
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            ZRechazado = IIf(IsNull(rstIvaComp!Rechazado), "0", rstIvaComp!Rechazado)
            rstIvaComp.Close
        End If
        
        Rem If Val(WVector1.TextMatrix(iRow, 5)) <> 0 And Val(WVector1.TextMatrix(iRow, 1)) <> 0 And ZRechazado = 0 Then
        If Val(WVector1.TextMatrix(iRow, 5)) <> 0 Then
        If (Val(WVector1.TextMatrix(iRow, 1)) <> 0 And ZRechazado = 0) Or Tipo2.Value = True Then
            
            
            Select Case Val(Left$(WVector1.TextMatrix(iRow, 0), 1))
                Case 1
                    XImpre1 = "Factura"
                Case 2
                    XImpre1 = "N.Debito"
                Case 3
                    XImpre1 = "N.Credito"
                Case 5
                    XImpre1 = "Anticipo"
                Case 99
                    XImpre1 = "Varios"
                Case Else
                    XImpre1 = ""
            End Select
                                
            XImpre2 = Left$(WVector1.TextMatrix(iRow, 4), 8)
                
            Rem spIvacomp = "ConsultaIvacomp " + "'" + XImpre2 + "'"
            Rem Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstIvacomp.RecordCount > 0 Then
            Rem     XImpre2 = rstIvacomp!Numero
            Rem     rstIvacomp.Close
            Rem End If
                    
            WTipo = WVector1.TextMatrix(iRow, 1)
            WLetra = WVector1.TextMatrix(iRow, 2)
            WPunto = WVector1.TextMatrix(iRow, 3)
            WNumero = WVector1.TextMatrix(iRow, 4)
                    
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XImpre3 = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                XImpre3 = ""
            End If
                        
            WImpre4 = Val(WVector1.TextMatrix(iRow, 5))
            Rem If Val(WTipo) = 3 Or Val(WTipo) = 5 Then
            Rem    WImpre4 = WImpre4 * -1
            Rem End If
            If WTipoiva = 2 Then
                WImpre4 = WImpre4 / 1.21
            End If
            Call Redondeo(WImpre4)
            XImpre4 = Str$(WImpre4)
            
            
            If WPorceIbCaba <> 0 Then
                WRete = Val(XImpre4) * (WPorceIbCaba / 100)
                    Else
                If WTipoIb = 3 Then
                    WRete = Val(XImpre4) * (2.5 / 100)
                        Else
                    WRete = Val(XImpre4) * (4.5 / 100)
                End If
            End If
            Call Redondeo(WRete)
            WImpoRetenido = WImpoRetenido + WRete
                            
            WRenglon = WRenglon + 1
            With rstImpreRetIb
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "1" + Auxi + Auxi1
                !Tipo = 1
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIbCiudad)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(DA)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                If WPorceIbCaba <> 0 Then
                    !Categoria1 = "SUJETO A RETENCION " + Str$(WPorceIbCaba) + "%"
                        Else
                    If WTipoIb = 3 Then
                        !Categoria1 = "SUJETO A RETENCION 2.50%"
                            Else
                        !Categoria1 = "SUJETO A RETENCION 4.50%"
                    End If
                End If
                !Importe1 = Val(XImpre4)
                If WPorceIbCaba <> 0 Then
                    !Porce1 = WPorceIbCaba
                        Else
                    If WTipoIb = 3 Then
                        !Porce1 = 2.5
                            Else
                        !Porce1 = 4.5
                    End If
                End If
                !Retencion1 = WRete
                .Update
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "2" + Auxi + Auxi1
                !Tipo = 2
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIbCiudad)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(DA)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                If WPorceIbCaba <> 0 Then
                    !Categoria1 = "SUJETO A RETENCION " + Str$(WPorceIbCaba) + "%"
                        Else
                    If WTipoIb = 3 Then
                        !Categoria1 = "SUJETO A RETENCION 2.5%"
                            Else
                        !Categoria1 = "SUJETO A RETENCION 4.50%"
                    End If
                End If
                !Importe1 = Val(XImpre4)
                If WPorceIbCaba <> 0 Then
                    !Porce1 = WPorceIbCaba
                        Else
                    If WTipoIb = 3 Then
                        !Porce1 = 2.5
                            Else
                        !Porce1 = 4.5
                    End If
                End If
                !Retencion1 = WRete
                .Update
            End With
        End If
        End If
        
    Next iRow
    
    For Ciclo = WRenglon + 1 To 12
        With rstImpreRetIb
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIbCiudad)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIbCiudad)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
        End With
    Next Ciclo
    
    LISTADO.GroupSelectionFormula = ""
    LISTADO.SelectionFormula = ""
        
    LISTADO.ReportFileName = "ImpreretibCiudad.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub



Private Sub ImpreretIva()

    On Error GoTo WError
        
    WRenglon = 0
    DA = 0
    With rstImpreRetIb
        .Index = "Orden"
        .Seek ">=", 0
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

    Select Case Val(WEmpresa)
        Case 1, 10
            WEmpNombre = "SURFACTAN S.A."
            WEmpDireccion = "Malvinas Argentinas 4589"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-54916508-3"
            WNroIb = "902-913585-2"
            WNroAgente = ""
        Case Else
            WEmpNombre = "PELLITAL S.A."
            WEmpDireccion = "Uruguay 2671"
            WEmpLocalidad = "1644 Victoria Bs.As. Argentina"
            WEmpCuit = "30-61052459-8"
            WNroIb = ""
            WNroAgente = ""
    End Select
    
    ImpreCopia(1) = "Original"
    ImpreCopia(2) = "Duplicado"
    
    WReteIva = 0
                
    For iRow = 1 To 12
            
        WRow = iRow
                
        XTipo = Left$(WVector1.TextMatrix(iRow, 1), 2)
        Select Case Val(Left$(WVector1.TextMatrix(iRow, 1), 2))
            Case 1
                XImpre1 = "Factura"
            Case 2
                XImpre1 = "N.Debito"
            Case 3
                XImpre1 = "N.Credito"
            Case 5
                XImpre1 = "Anticipo"
            Case 99
                XImpre1 = "Varios"
            Case Else
                XImpre1 = ""
        End Select
        
        XLetra = Left$(WVector1.TextMatrix(iRow, 2), 1)
        XPunto = Left$(WVector1.TextMatrix(iRow, 3), 4)
        XNumero = Left$(WVector1.TextMatrix(iRow, 4), 8)
        XImpre2 = Left$(WVector1.TextMatrix(iRow, 4), 8)
        XImporte = WVector1.TextMatrix(iRow, 5)
                
        If Val(XImporte) <> 0 Then
                
            XBruto = Val(XImporte)
            XNeto = (XBruto / 1.21)
            XIva = XBruto - XNeto

            If XNeto >= 1000 And Val(XNumero) <> 0 Then
            
                Sql1 = "Select *"
                Sql2 = " FROM IvaComp"
                Sql3 = " Where IvaComp.Proveedor = " + "'" + Proveedor.Text + "'"
                Sql4 = " and IvaComp.Tipo = " + "'" + XTipo + "'"
                Sql5 = " and IvaComp.Letra = " + "'" + XLetra + "'"
                Sql6 = " and IvaComp.Punto = " + "'" + XPunto + "'"
                Sql7 = " and IvaComp.Numero = " + "'" + XNumero + "'"
                spIvaComp = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    WReteIva = rstIvaComp!Iva21
                    rstIvaComp.Close
                End If
                
            End If
                    
            WTipo = XTipo
            WLetra = XLetra
            WPunto = XPunto
            WNumero = XNumero
                    
            ClaveCtaprv = Proveedor.Text + WLetra + WTipo + WPunto + WNumero
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                XImpre3 = RstCtaPrv!Fecha
                RstCtaPrv.Close
                    Else
                XImpre3 = ""
            End If
                        
            WImpre4 = Val(XImporte)
                            
            WRenglon = WRenglon + 1
            With rstImpreRetIb
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "1" + Auxi + Auxi1
                !Tipo = 1
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(DA)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                !Categoria1 = "SUJETO A RETENCION 0.75%"
                !Importe1 = Val(XImpre4)
                !Porce1 = WReteIva
                !Retencion1 = WReteIva
                .Update
                .AddNew
                Auxi = Orden.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                !Clave = "2" + Auxi + Auxi1
                !Tipo = 2
                !Orden = Val(Orden.Text)
                !Renglon = WRenglon
                !Empresa = WEmpNombre
                !Direccion = WEmpDireccion
                !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
                !Localidad = WEmpLocalidad
                !Fecha = Fecha.Text
                !Cuit = WEmpCuit
                !Copia = ImpreCopia(DA)
                !NroIb = WNroIb
                !NroAgente = WNroAgente
                !NombrePrv = DesProveedor.Caption
                !DireccionPrv = WPrvDireccion
                !CuitPrv = WPrvCuit
                !NroIbPrv = WPrvIb
                !Tipo1 = XImpre1
                !Numero1 = XImpre2
                !Fecha1 = XImpre3
                !Categoria1 = ""
                !Importe1 = Val(XImpre4)
                !Porce1 = WReteIva
                !Retencion1 = WReteIva
                .Update
            End With
                    
        End If
    Next iRow
    
    For Ciclo = WRenglon + 1 To 12
        With rstImpreRetIb
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "1" + Auxi + Auxi1
            !Tipo = 1
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
            .AddNew
            Auxi = Orden.Text
            Call Ceros(Auxi, 6)
            Auxi1 = Ciclo
            Call Ceros(Auxi1, 2)
            !Clave = "2" + Auxi + Auxi1
            !Tipo = 2
            !Orden = Val(Orden.Text)
            !Renglon = XCiclo
            !Empresa = WEmpNombre
            !Direccion = WEmpDireccion
            !Titulo = "Nro.Certificado  : " + Str$(WCertificadoIva)
            !Localidad = WEmpLocalidad
            !Fecha = Fecha.Text
            !Cuit = WEmpCuit
            !Copia = ImpreCopia(DA)
            !NroIb = WNroIb
            !NroAgente = WNroAgente
            !NombrePrv = DesProveedor.Caption
            !DireccionPrv = WPrvDireccion
            !CuitPrv = WPrvCuit
            !NroIbPrv = WPrvIb
            !Tipo1 = ""
            !Numero1 = ""
            !Fecha1 = ""
            !Categoria1 = ""
            !Importe1 = 0
            !Porce1 = 0
            !Retencion1 = 0
            .Update
        End With
    Next Ciclo
        
    LISTADO.ReportFileName = "Impreretiva.rpt"
    LISTADO.Destination = 1
    LISTADO.DataFiles(0) = WEmpresa + "Auxi.mdb"
    LISTADO.CopiesToPrinter = 1
    LISTADO.Action = 1
        
    Exit Sub
        
WError:
    Resume Next

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        If Ayuda.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Nombre LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Nombre"
            spProveedor = ZSql
                Else
            spProveedor = "ListaProveedoresOrdConsulta"
        End If
    
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
        End If
    
    End If

End Sub

Private Sub CargaCarpeta_Click()

    IngreCarpeta.Height = 3375
    IngreCarpeta.Left = 4080
    IngreCarpeta.Top = 2160
    IngreCarpeta.Width = 3015
        
   IngreCarpeta.Visible = True
        
    Carpeta.SetFocus
    
End Sub

Private Sub GrabaCarpeta_Click()
    
    ZCarpeta(1) = Carpeta.Text
    ZCarpeta(2) = Carpeta1.Text
    ZCarpeta(3) = Carpeta2.Text
    ZCarpeta(4) = Carpeta3.Text
    ZCarpeta(5) = Carpeta4.Text
    ZZProveedor = ""
            
    For Ciclo = 1 To 5
            
        If Val(ZCarpeta(Ciclo)) <> 0 Then
                
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                    
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + ZCarpeta(Ciclo) + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        ZZProveedor = rstOrden!Proveedor
                        rstOrden.Close
                        WEntra = "S"
                    End If
                            
                    If WEntra = "S" Then
                        If Proveedor.Text <> "10167878480" And Proveedor.Text <> "10000000100" And Proveedor.Text <> "10071081483" Then
                            If Proveedor.Text <> ZZProveedor Then
                                m$ = "El Proveedor de la Carpeta no coincide con el de la orden de pago"
                                A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                                Exit Sub
                            End If
                        End If
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
   
         If WEntra = "N" Then
         
               m$ = "Numero de carpeta incorrecto"
                A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                Exit Sub
            End If
            
                    
       End If
            
    Next Ciclo

    IngreCarpeta.Visible = False
 End Sub

 Private Sub Carpeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
           Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                   ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta.Text + "'"
                    spOrden = ZSql
                   Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                       rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta1.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta1.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta1.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta2.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta2.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta2.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta3.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta3.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta3.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta4.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Carpeta4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Carpeta4.Text) <> 0 Then
        
            XEmpresa = WEmpresa
            WEntra = "N"
        
            Select Case Val(XEmpresa)
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
                    
            For Cicla = 1 To ZHasta
                If CargaEmpresa(Cicla, 1) <> "" Then
                
                    WEmpresa = CargaEmpresa(Cicla, 1)
                    txtOdbc = CargaEmpresa(Cicla, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden.Carpeta = " + "'" + Carpeta4.Text + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                        WEntra = "S"
                        Exit For
                    End If
                    
                End If
            Next Cicla
    
            Call Conecta_Empresa
            
            If WEntra = "S" Then
                Carpeta.SetFocus
            End If
    
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Alta_Dife()

    XNroInterno = ""
    spIvaComp = "ListaIvacompNumero"
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveLast
            XNroInterno = Str$(rstIvaComp!NroInterno + 1)
        End With
        rstIvaComp.Close
    End If
    WNumeroDife = XNroInterno
    
    Call Ceros(XNroInterno, 6)
    Call Ceros(WTipoDife, 2)
    Call Ceros(WPuntoDife, 4)
    Call Ceros(WNumeroDife, 8)
    
    Rem graba el iva compras
    
    XProveedor = Proveedor.Text
    XTipo = WTipoDife
    XLetra = WLetraDife
    XPunto = WPuntoDife
    XNumero = WNumeroDife
    XFecha = Fecha.Text
    Xvencimiento = Fecha.Text
    XVencimiento1 = Fecha.Text
    XPeriodo = Fecha.Text
    XImpoNeto = Str$(WNetoDife)
    XIva21 = Str$(WIvaDife)
    XIva5 = ""
    XIva27 = ""
    XIb = ""
    XExento = ""
    Select Case Val(WTipoDife)
        Case 1
            XImpre = "FC"
        Case 2
            XImpre = "ND"
        Case 3
            XImpre = "NC"
        Case Else
            XImpre = "  "
    End Select
    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XContado = "2"
    XEmpresa = "1"
    XNetolist = ""
    XExentolist = ""
    XParidad = ""
    XPAgo = "1"
    
    XParam = "'" + XNroInterno + "','" _
                 + XProveedor + "','" + XTipo + "','" _
                 + XLetra + "','" _
                 + XPunto + "','" + XNumero + "','" _
                 + XFecha + "','" _
                 + Xvencimiento + "','" _
                 + XVencimiento1 + "','" + XPeriodo + "','" _
                 + XImpoNeto + "','" _
                 + XIva21 + "','" _
                 + XIva5 + "','" + XIva27 + "','" _
                 + XIb + "','" + XExento + "','" _
                 + XContado + "','" _
                 + XImpre + "','" + XOrdFecha + "','" _
                 + XEmpresa + "','" + XNetolist + "','" _
                 + XExentolist + "','" _
                 + XParidad + "','" _
                 + XPAgo + "'"
                
    spIvaComp = "AltaIvaCompras " + XParam
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
                
                
    Rem graba las imputaciones contables
                                        
    RenglonDife = 0
    
    
    Rem renglon nro 1
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "6107"
            XDebito = Str$(Abs(WNetoDife))
            XCredito = ""
        Case Else
            XCuenta = "7308"
            XDebito = ""
            XCredito = Str$(Abs(WNetoDife))
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem renglon nro 2
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "151"
            XDebito = Str$(Abs(WIvaDife))
            XCredito = ""
        Case Else
            XCuenta = "151"
            XDebito = ""
            XCredito = Str$(Abs(WIvaDife))
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem renglon nro 3
                                        
    RenglonDife = RenglonDife + 1
    Auxi1 = Str$(RenglonDife)
    Call Ceros(Auxi1, 2)
    XRenglon = Auxi1
                        
    XTipomovi = "2"
    XTipocomp = WTipoDife
    XLetracomp = WLetraDife
    XPuntocomp = WPuntoDife
    XNrocomp = WNumeroDife
    XFecha = Fecha.Text
    XObservaciones = ""
    Select Case Val(WTipoDife)
        Case 2
            XCuenta = "2001"
            XDebito = ""
            XCredito = Str$(Abs(WNetoDife) + Abs(WIvaDife))
        Case Else
            XCuenta = "2001"
            XDebito = Str$(Abs(WNetoDife) + Abs(WIvaDife))
            XCredito = ""
    End Select
    XFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XTitulo = "Compras"
    XEmpresa = "1"
    XClave = XTipomovi + XNroInterno + XRenglon
    XDebitolist = ""
    XCreditolist = ""
    
    XParam = "'" + XClave + "','" _
                 + XTipomovi + "','" + XProveedor + "','" _
                 + XTipocomp + "','" _
                 + XLetracomp + "','" + XPuntocomp + "','" _
                 + XNrocomp + "','" _
                 + XRenglon + "','" _
                 + XFecha + "','" + XObservaciones + "','" _
                 + XCuenta + "','" _
                 + XDebito + "','" _
                 + XCredito + "','" + XFechaOrd + "','" _
                 + XTitulo + "','" + XEmpresa + "','" _
                 + XDebitolist + "','" _
                 + XCreditolist + "','" _
                 + XNroInterno + "'"
                                
    spImputac = "AltaImputacion " + XParam
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    Rem alta en la cuenta corriente de proveedores
                
    XProveedor = Proveedor.Text
    XLetra = WLetraDife
    XTipo = WTipoDife
    XPunto = WPuntoDife
    XNumero = WNumeroDife
    XFecha = Fecha.Text
    XEstado = "1"
    Xvencimiento = Fecha.Text
    XVencimiento1 = Fecha.Text
    XNroInterno = XNroInterno
    XTotal = Str$(WNetoDife + WIvaDife)
    XSaldo = Str$(WNetoDife + WIvaDife)
    XClave = Proveedor.Text + WLetraDife + WTipoDife + WPuntoDife + WNumeroDife
    XOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    XOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Select Case Val(WTipoDife)
        Case 1
            XImpre = "FC"
        Case 2
            XImpre = "ND"
        Case 3
            XImpre = "NC"
        Case Else
            XImpre = ""
    End Select
    XEmpresa = "1"
    XSaldolist = ""
    Xlista = ""
    XAcumulado = ""
    XParidad = ""
    XPAgo = "1"
                    
    XParam = "'" + XClave + "','" _
                 + XProveedor + "','" + XLetra + "','" _
                 + XTipo + "','" _
                 + XPunto + "','" + XNumero + "','" _
                 + XFecha + "','" _
                 + XEstado + "','" _
                 + Xvencimiento + "','" + XVencimiento1 + "','" _
                 + XTotal + "','" _
                 + XSaldo + "','" _
                 + XOrdFecha + "','" + XOrdVencimiento + "','" _
                 + XImpre + "','" + XEmpresa + "','" _
                 + XSaldolist + "','" _
                 + XNroInterno + "','" + Xlista + "','" _
                 + XAcumulado + "','" _
                 + XParidad + "','" _
                 + XPAgo + "'"
                    
    spConsulta = "AltaCtaPrv " + XParam
    Set rstConsulta = db.OpenRecordset(spConsulta + cParam, dbOpenSnapshot, dbSQLPassThrough)
    
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
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
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
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
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
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            End If
        End If
        Rem Call Calcula_Click
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

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

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
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""
    
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

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
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
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            If WControlII <> "N" Then
                Call StartEdit
            End If
            WControlII = ""

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
        Case 6
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case 12
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 7
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
            If Tipo1.Value = True Then
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                        Else
                    WControl = "N"
                End If
                    Else
                If Val(WVector1.Text) = 0 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                    WVector1.Col = 4
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 2
            WVector1.Text = Left$(WVector1.Text, 1)
            If Tipo1.Value = True Then
                If WVector1.Text = "A" Or WVector1.Text = "C" Or WVector1.Text = "X" Or WVector1.Text = "E" Then
                    WControl = "S"
                        Else
                    WControl = "N"
                End If
            End If
                
        Case 3
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 4)
            WVector1.Text = Auxi$
            
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
                    
            ClaveCtaprv = Proveedor.Text
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 2)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 1)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 3)
            ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 4)
            
            spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
            Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
            If RstCtaPrv.RecordCount > 0 Then
                If Val(WVector1.TextMatrix(WVector1.Row, 5)) = 0 Then
                    WVector1.TextMatrix(WVector1.Row, 5) = RstCtaPrv!Saldo
                    RstCtaPrv.Close
                    Call Suma_Datos
                End If
                Rem WVector1.Col = 4
                    Else
                WControl = "N"
            End If
            
        Case 5
            
            ZZZValida = "S"
            ZZZLetra = WVector1.TextMatrix(WVector1.Row, 2)
            ZZZTipo = WVector1.TextMatrix(WVector1.Row, 1)
            ZZZPunto = WVector1.TextMatrix(WVector1.Row, 3)
            ZZZNumero = WVector1.TextMatrix(WVector1.Row, 4)
            If Trim(ZZZLetra) = "" And Val(ZZZTipo) = 0 And Val(ZZZPunto) = 0 And Val(ZZZNumero) = 0 Then
                ZZZValida = "N"
            End If
            
            If Tipo1.Value = True And ZZZValida = "S" Then
            
                ClaveCtaprv = Proveedor.Text
                ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 2)
                ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 1)
                ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 3)
                ClaveCtaprv = ClaveCtaprv + WVector1.TextMatrix(WVector1.Row, 4)
                        
                spCtaprv = "ConsultaCtaprv " + "'" + ClaveCtaprv + "'"
                Set RstCtaPrv = db.OpenRecordset(spCtaprv, dbOpenSnapshot, dbSQLPassThrough)
                If RstCtaPrv.RecordCount > 0 Then
                    Saldo = RstCtaPrv!Saldo
                    XNroInterno = Str$(RstCtaPrv!NroInterno)
                    XLetra = RstCtaPrv!Letra
                    XPunto = RstCtaPrv!Punto
                    RstCtaPrv.Close
                        Else
                    Saldo = 0
                End If
                
                XSaldo = Val(WVector1.TextMatrix(WVector1.Row, 5))
                If XSaldo > Saldo Then
                        
                    XSaldo = 0
                    WVector1.TextMatrix(WVector1.Row, 5) = ""
                    WControl = "N"
                            
                        Else
                                
                    WVector1.TextMatrix(WVector1.Row, 5) = Pusing("#,###,###.##", WVector1.TextMatrix(WVector1.Row, 5))
                    
                    XRow = WVector1.Row
                    Call Suma_Datos
                        
                    spIvaComp = "Consultaivacomp " + "'" + XNroInterno + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        XParidad = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
                        XPAgo = IIf(IsNull(rstIvaComp!Pago), "0", rstIvaComp!Pago)
                        XLetra = rstIvaComp!Letra
                        XPunto = rstIvaComp!Punto
                        rstIvaComp.Close
                    End If
            
                    ParidadTotal = 0
                    If XPAgo = 2 Then
                        spCambio = "ConsultaCambio " + "'" + FechaParidad.Text + "'"
                        Set rstCambio = db.OpenRecordset(spCambio, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCambio.RecordCount > 0 Then
                            ParidadTotal = rstCambio!Cambio
                            rstCambio.Close
                        End If
                        WSaldo = XSaldo
                        WSaldoUs = (XSaldo / XParidad) * ParidadTotal
                        WDife = WSaldoUs - WSaldo
                        Call Redondeo(WDife)
                
                        If WDife <> 0 Then
                            If WDife > 0 Then
                
                                XRow = XRow + 1
                                
                                WVector1.Row = XRow
                                
                                WVector1.Col = 1
                                WVector1.Text = "02"
                            
                                WVector1.Col = 2
                                WVector1.Text = XLetra
                            
                                WVector1.Col = 3
                                WVector1.Text = XPunto
                    
                                WVector1.Col = 4
                                WVector1.Text = "99999999"
                    
                                WVector1.Col = 5
                                WVector1.Text = WDife
                                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                                WVector1.Col = 6
                                WVector1.Text = "N/D por Diferencia de Cambio "
                    
                                    Else
                
                                XRow = XRow + 1
                        
                                WVector1.Row = XRow
                                
                                WVector1.Col = 1
                                WVector1.Text = "03"
                
                                WVector1.Col = 2
                                WVector1.Text = XLetra
                
                                WVector1.Col = 3
                                WVector1.Text = XPunto
                
                                WVector1.Col = 4
                                WVector1.Text = "99999999"
                
                                WVector1.Col = 5
                                WVector1.Text = WDife
                                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                                WVector1.Col = 6
                                WVector1.Text = "N/C por Diferencia de Cambio "
                    
                            End If
                        End If
                        
                    End If
                    
                End If
                
                    Else
                    
                columna = WVector1.Row
                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                Call Suma_Datos
                
            End If
            
        Case 6
            ZZZValida = "N"
            ZZZLetra = WVector1.TextMatrix(WVector1.Row, 2)
            ZZZTipo = WVector1.TextMatrix(WVector1.Row, 1)
            ZZZPunto = WVector1.TextMatrix(WVector1.Row, 3)
            ZZZNumero = WVector1.TextMatrix(WVector1.Row, 4)
            If Trim(ZZZLetra) = "" And Val(ZZZTipo) = 0 And Val(ZZZPunto) = 0 And Val(ZZZNumero) = 0 Then
                Rem ZZZValida = "S"
            End If
            
            If Tipo3.Value = True Or ZZZValida = "S" Then
                WProceso = 0
                Cuenta.Text = WCuenta(WVector1.Row, 1)
                IngreCuenta.Visible = True
                WControl = "N"
                WControlII = "N"
                Cuenta.SetFocus
                    Else
                If Tipo4.Value = True Then
                    spBanco = "ConsultaBanco " + "'" + Banco.Text + "'"
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                        WCuenta(WVector1.Row, 1) = rstBanco!Cuenta
                        rstBanco.Close
                            Else
                        WCuenta(WVector1.Row, 1) = "999999"
                    End If
                End If
                If Tipo5.Value = True Then
                    WCuenta(WVector1.Row, 1) = "111"
                End If
            End If
            
        Case 7
            If Len(WVector1.Text) = 31 Then
                Lectora.Text = WVector1.Text
                WVector1.Text = "99"
            End If
            
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Or Val(WVector1.Text) = 5 Or Val(WVector1.Text) = 6 Or Val(WVector1.Text) = 7 Or Val(WVector1.Text) = 8 Or Val(WVector1.Text) = 99 Then
                
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 2)
                WVector1.Text = Auxi$
                
                        
                Select Case Val(WVector1.Text)
                    Case 1
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                        WVector1.Col = 10
                        WVector1.Text = ""
                        WVector1.Col = 11
                        WVector1.Text = "Efectivo"
                                
                    Case 5
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                        WVector1.Col = 10
                        WVector1.Text = ""
                        WVector1.Col = 11
                        WVector1.Text = "U$S"
                                
                    Case 3
                        WVector1.Col = 7
                        WVector1.Text = ""
                        Rem Call Consulta_Click
                        WControl = "N"
                        
                        Opcion.Clear
                        Opcion.AddItem "Proveedores"
                        Opcion.AddItem "Cuenta Corrientes"
                        Opcion.AddItem "Cheques terceros"
                        Rem Opcion.Visible = True
                        Opcion.ListIndex = 2
    
                        Rem Call Opcion_Click
                        
                    Case 4
                        WVector1.Col = 7
                        WVector1.Text = ""
                        Rem Call Consulta_Click
                        WControl = "N"
                        
                        Opcion.Clear
                        Opcion.AddItem "Proveedores"
                        Opcion.AddItem "Cuenta Corrientes"
                        Opcion.AddItem "Cheques terceros"
                        Opcion.AddItem "Docuementos"
                        Rem Opcion.Visible = True
                        Opcion.ListIndex = 3
    
                        Rem Call Opcion_Click
                                
                    Case 6
                        WProceso = 1
                        Cuenta.Text = WCuenta(WVector1.Row, 2)
                        IngreCuenta.Visible = True
                        Cuenta.SetFocus
                        WControl = "N"
                        WControlII = "N"
                                
                    Case 7
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                        WVector1.Col = 10
                        WVector1.Text = ""
                        WVector1.Col = 11
                        WVector1.Text = "Patacones"
                                
                    Case 8
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                        WVector1.Col = 10
                        WVector1.Text = ""
                        WVector1.Col = 11
                        WVector1.Text = "Lecop"
                        
                    Case 99
                        WVector1.Text = ""
                        WControl = "N"
                        WControlII = "N"
                        Lectora.Visible = True
                        Call Lectora_Keypress(13)
                        
                    Case Else
                                
                End Select
                        
                    Else
                    
                WControl = "N"
                       
            End If
        
        Case 8
            WVector1.Col = 7
            If Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Then
                WVector1.Col = 8
                WControl = "N"
                    Else
                WVector1.Col = 8
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 8)
                WVector1.Text = Auxi$
            End If
                
        Case 9
            If Len(WVector1.Text) = 5 Then
                If Right$(WVector1.Text, 2) < 6 Then
                    WVector1.Text = WVector1.Text + "/2013"
                        Else
                    WVector1.Text = WVector1.Text + Right$(Fecha.Text, 5)
                End If
            End If
            Call Valida_fecha1(WVector1.Text, Auxi)
            If Auxi <> "S" Then
                WControl = "N"
                WControl = "N"
                    Else
                If Val(WVector1.TextMatrix(WVector1.Row, 10)) <> 0 Then
                    WVector1.Col = WVector1.Col + 2
                End If
            End If
                
        Case 10
            ClaveBanco = WVector1.Text
            spBanco = "ConsultaBanco " + "'" + ClaveBanco + "'"
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                WVector1.Col = 11
                WVector1.Text = rstBanco!Nombre
                rstBanco.Close
                    Else
                WControl = "N"
            End If

        Case 12
            If Val(WVector1.Text) > 0 Then
                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
            End If
            Call Suma_Datos
            
        Case Else
    End Select
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
    WVector1.Cols = 15
    WVector1.FixedRows = 1
    WVector1.Rows = 14
    
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
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 6
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1850
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 40
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 12
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###,###.##"
            Case 13
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 20
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 14
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 20
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 3
        WVector1.Col = Ciclo
        WTituloVector(Ciclo).Text = WVector1.Text
        WTituloVector(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTituloVector(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTituloVector(Ciclo).Width = WVector1.CellWidth
        WTituloVector(Ciclo).Height = WVector1.CellHeight
        WTituloVector(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = 11400
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

Private Sub Lectora_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Lectora.Visible = False
        If Len(Lectora.Text) = 31 Then
        
                Entra = "N"
                For iRow = 1 To 12
                    If Val(WVector1.TextMatrix(iRow, 7)) = 0 Then
                        XRow = iRow
                        Entra = "S"
                        Exit For
                    End If
                Next iRow
                
                If Entra = "N" Then
                    m$ = "La cantidad de valores entregados supera los 12"
                    A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                    Exit Sub
                End If
                
                ZZBanco = Mid$(Lectora, 2, 3)
                ZZSucursal = Mid$(Lectora, 5, 3)
                ZZNroCheque = Mid$(Lectora, 12, 8)
                ZZNroCuenta = Mid$(Lectora, 20, 11)
                
                For ZZCiclo = 1 To 12
                    If WVector1.TextMatrix(ZZCiclo, 8) = ZZNroCheque Then
                        m$ = "Cheque ya cargado"
                        A% = MsgBox(m$, 0, "Emision de Ordenes de Pago")
                        Exit Sub
                    End If
                Next ZZCiclo
                    
                Rem VER ACA NUMERO DE CUENTA
                If ZZNroCuenta = "00030521158" Then
                                                
                    WVector1.Row = XRow
                    
                    WVector1.Col = 7
                    WVector1.Text = "02"
                    
                    WVector1.Col = 8
                    WVector1.Text = ZZNroCheque
                    
                    WVector1.Col = 9
                    WVector1.Text = ""
                    
                    WVector1.Col = 10
                    WVector1.Text = "3"
                    
                    WVector1.Col = 11
                    WVector1.Text = "NACION (SF)"
                    
                    WVector1.Col = 12
                    WVector1.Text = ""
                    
                    WVector1.Col = 13
                    WVector1.Text = ""
                    
                    WVector1.Col = 14
                    WVector1.Text = ""
                    
                    WVector1.Col = 9
                    Call StartEdit
                
                
                        Else
                        
                    If ZZNroCuenta = "00830029723" Then
                
                        WVector1.Row = XRow
                    
                        WVector1.Col = 7
                        WVector1.Text = "02"
                    
                        WVector1.Col = 8
                        WVector1.Text = ZZNroCheque
                    
                        WVector1.Col = 9
                        WVector1.Text = ""
                    
                        WVector1.Col = 10
                        WVector1.Text = "8"
                    
                        WVector1.Col = 11
                        WVector1.Text = "HSBC (SF"
                    
                        WVector1.Col = 12
                        WVector1.Text = ""
                    
                        WVector1.Col = 13
                        WVector1.Text = ""
                    
                        WVector1.Col = 14
                        WVector1.Text = ""
                    
                        WVector1.Col = 9
                        Call StartEdit
                
                            Else
                            
                        Entra = "S"
                        
                        Sql1 = "Select *"
                        Sql2 = " FROM Recibos"
                        Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
                        spRecibos = Sql1 + Sql2 + Sql3
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        If rstRecibos.RecordCount > 0 Then
                        
                            Entra = "N"
                            If rstRecibos!Estado2 = "P" Then
                
                                WVector1.Row = XRow
                        
                                WVector1.Col = 7
                                WVector1.Text = "3"
                        
                                WVector1.Col = 8
                                WVector1.Text = rstRecibos!Numero2
                        
                                WVector1.Col = 9
                                WVector1.Text = rstRecibos!Fecha2
                        
                                WVector1.Col = 10
                                WVector1.Text = ""
                        
                                Busqueda.Text = rstRecibos!Banco2
                                Foundpos = Busqueda.Find("/")
                        
                                If Foundpos > 0 Then
                                    WVector1.Col = 11
                                    WVector1.Text = Left$(rstRecibos!Banco2, Foundpos)
                                        Else
                                    WVector1.Col = 11
                                    WVector1.Text = rstRecibos!Banco2
                                End If
                        
                        
                                WVector1.Col = 12
                                WVector1.Text = rstRecibos!Importe2
                                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                        
                                WVector1.Col = 13
                                WVector1.Text = "1" + rstRecibos!Clave
                        
                                WVector1.Col = 14
                                WVector1.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
                        
                                WVector1.Col = 7
                                WVector1.Text = "3"
                        
                                rstRecibos.Close
                        
                                Call Suma_Datos
                        
                                Pantalla.List(Indice) = ""
                            
                                If XRow < 12 Then
                                    XRow = XRow + 1
                                End If
                    
                                WVector1.Row = XRow
                                WVector1.Col = 7
                    
                                Call StartEdit
                            
                            End If
                            
                        End If
                        
                        If Entra = "S" Then
            
                            Sql1 = "Select *"
                            Sql2 = " FROM RecibosProvi"
                            Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
                            Sql4 = " and RecibosProvi.Estado2 = " + "'" + "P" + "'"
                            spRecibosProvi = Sql1 + Sql2 + Sql3 + Sql4
                            Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
                            If rstRecibosProvi.RecordCount > 0 Then
                            
                                WVector1.Row = XRow
                                Entra = "N"
                    
                                WVector1.Col = 7
                                WVector1.Text = "3"
                    
                                WVector1.Col = 8
                                WVector1.Text = rstRecibosProvi!Numero2
                    
                                WVector1.Col = 9
                                WVector1.Text = rstRecibosProvi!Fecha2
                    
                                WVector1.Col = 10
                                WVector1.Text = ""
                    
                                Busqueda.Text = rstRecibosProvi!Banco2
                                Foundpos = Busqueda.Find("/")
                    
                                If Foundpos > 0 Then
                                    WVector1.Col = 11
                                    WVector1.Text = Left$(rstRecibosProvi!Banco2, Foundpos)
                                        Else
                                    WVector1.Col = 11
                                    WVector1.Text = rstRecibosProvi!Banco2
                                End If
                    
                                WVector1.Col = 12
                                WVector1.Text = rstRecibosProvi!Importe2
                                WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                                WVector1.Col = 13
                                WVector1.Text = "2" + rstRecibosProvi!Clave
                    
                                WVector1.Col = 14
                                WVector1.Text = IIf(IsNull(rstRecibosProvi!Cuit), "", rstRecibosProvi!Cuit)
                    
                                WVector1.Col = 7
                                WVector1.Text = "3"
                    
                                rstRecibosProvi.Close
                    
                                Call Suma_Datos
                    
                                Pantalla.List(Indice) = ""
                        
                                If XRow < 12 Then
                                    XRow = XRow + 1
                                End If
                
                                WVector1.Row = XRow
                                WVector1.Col = 7
                
                                Call StartEdit
                            
                            End If
                            
                        End If
                    
                    End If
                        
                
                End If
                
            
        End If
        
        Lectora.Visible = False
    End If
    If KeyAscii = 27 Then
        Lectora.Visible = False
    End If
End Sub











Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
    
            Case 0
                If KeyCode = 13 Then
                    If Val(DBGrid1.Text) = 1 Or Val(DBGrid1.Text) = 2 Or Val(DBGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DBGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DBGrid1.Text = Auxi$
                        DBGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DBGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If
                
            Case 1
                Rem If KeyCode = 13 Then
                Rem     DBGrid1.Text = Left$(DBGrid1.Text, 1)
                Rem     If DBGrid1.Text = "A" Or DBGrid1.Text = "C" Then
                Rem         DBGrid1.Col = 2
                Rem         KeyCode = 0
                Rem         Rem no hago anda
                Rem             Else
                Rem         DBGrid1.Col = 1
                Rem         KeyCode = 0
                Rem     End If
                Rem End If
                
            Case 2
                Rem If KeyCode = 13 Then
                Rem     Auxi$ = Str$(Val(DBGrid1.Text))
                Rem     Call Ceros(Auxi$, 4)
                Rem     DBGrid1.Text = Auxi$
                Rem     DBGrid1.Col = 3
                Rem     KeyCode = 0
                Rem End If
                
            Case 3
                If KeyCode = 13 Then
                
                    Auxi$ = Str$(Val(DBGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DBGrid1.Text = Auxi$
                    
                    With rstCtaCte
                    
                    
                        DBGrid1.Col = 0
                        XTipo = DBGrid1.Text
                        
                        DBGrid1.Col = 3
                        XNumero = DBGrid1.Text
                        
                        ClaveCtacte = XTipo + XNumero + "01"
                        
                        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCte.RecordCount > 0 Then
                        
                            DBGrid1.Col = 4
                            XRow = DBGrid1.Row
                            If Val(DBGrid1.Text) = 0 Then
                                DBGrid1.Text = !Saldo
                                Call Suma_Datos
                                DBGrid1.Col = 4
                                DBGrid1.Row = XRow
                            End If
                            DBGrid1.Col = 4
                            KeyCode = 0
                                Else
                            DBGrid1.Col = 0
                            KeyCode = 0
                            
                            rstCtaCte.Close
                            
                        End If
                    End With
                End If
                
            Case 4
                If KeyCode = 13 Then
                    With rstCtaCte
                        DBGrid1.Col = 0
                        XTipo = DBGrid1.Text
                        DBGrid1.Col = 3
                        XNumero = DBGrid1.Text
                        
                        ClaveCtacte = XTipo + XNumero + "01"
                        
                        spCtaCte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCte.RecordCount > 0 Then
                            Saldo = Alinea("###,###.##", Str$(rstCtaCte!Saldo))
                            rstCtaCte.Close
                                Else
                            Saldo = 0
                        End If
                    
                    End With
                
                    DBGrid1.Col = 4
                    If Abs(Val(DBGrid1.Text)) > Abs(Val(Saldo)) Then
                        DBGrid1.Text = ""
                        DBGrid1.Col = 4
                        KeyCode = 0
                            Else
                        DBGrid1.Text = Alinea("###,###.##", DBGrid1.Text)
                        Call Suma_Datos
                        If DBGrid1.Row < 10 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                                Else
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                    End If
                End If
                
            Case 5
                If KeyCode = 13 Then
                    If Len(DBGrid1.Text) = 31 Then
                        Lectora.Text = DBGrid1.Text
                        Call Lectora_Keypress(13)
                            Else
                        If Val(DBGrid1.Text) = 1 Or Val(DBGrid1.Text) = 2 Or Val(DBGrid1.Text) = 3 Or Val(DBGrid1.Text) = 4 Or Val(DBGrid1.Text) = 99 Then
                            Auxi$ = Str$(Val(DBGrid1.Text))
                            Call Ceros(Auxi$, 2)
                            DBGrid1.Text = Auxi$
                            Select Case Val(DBGrid1.Text)
                                Case 1, 4
                                    DBGrid1.Col = 6
                                    DBGrid1.Text = ""
                                    DBGrid1.Col = 7
                                    DBGrid1.Text = ""
                                    DBGrid1.Col = 8
                                    DBGrid1.Text = ""
                                    DBGrid1.Col = 9
                                    KeyCode = 0
                                Case Else
                                    DBGrid1.Col = 6
                                    KeyCode = 0
                            End Select
                                Else
                            DBGrid1.Col = 5
                            KeyCode = 0
                        End If
                    End If
                    
                End If
                
            Case 6
                If KeyCode = 13 Then
                    Auxi$ = Str$(Val(DBGrid1.Text))
                    Call Ceros(Auxi$, 8)
                    DBGrid1.Text = Auxi$
                    DBGrid1.Col = 7
                    KeyCode = 0
                End If
                
            Case 7
                If KeyCode = 13 Then
                
                    If Len(DBGrid1.Text) = 5 Then
                        If Val(Right$(DBGrid1.Text, 2)) > 6 Then
                            DBGrid1.Text = DBGrid1.Text + "/2013"
                                Else
                            DBGrid1.Text = DBGrid1.Text + "/2014"
                        End If
                    End If
                    
                    DBGrid1.Col = 7
                    Call Valida_fecha1(DBGrid1.Text, Auxi)
                    Rem Call Valida_fecha(DbGrid1.Text, Auxi)
                    
                    If Auxi <> "S" Then
                    
                        DBGrid1.Col = 7
                        KeyCode = 0
                        
                                Else
                                
                        ZPasa = ""
                        ZFecha = DBGrid1.Text
                        DBGrid1.Col = 5
                        ZTipo = Val(DBGrid1.Text)
        
                        WDias = 0
                        WFechaDesde = ZFecha
                        WFechaHasta = Fecha.Text
        
                        WOrdFechaDesde = Right$(WFechaDesde, 4) + Mid$(WFechaDesde, 4, 2) + Left$(WFechaDesde, 2)
                        WOrdFechaHasta = Right$(WFechaHasta, 4) + Mid$(WFechaHasta, 4, 2) + Left$(WFechaHasta, 2)
        
                        If ZTipo = 2 And WOrdFechaDesde < WOrdFechaHasta Then
        
                            Do
                                WDias = WDias + 1
                                XFec1 = WFechaDesde
                                SumaDia = 2
                                Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                                WFechaDesde = XFec2
                                If WFechaDesde = WFechaHasta Then
                                    Exit Do
                                End If
                            Loop
            
                            If WDias > 30 Then
                                ZPasa = "N"
                            End If
            
                        End If
                        
                        If ZPasa = "N" Then
                            m1$ = "Error en la carga de fecha de cheque"
                            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                            DBGrid1.Col = 7
                            KeyCode = 0
                                Else
                            DBGrid1.Col = 8
                            If Trim(DBGrid1.Text) <> "" Then
                                DBGrid1.Col = 9
                            End If
                            KeyCode = 0
                        End If
                    
                    End If
                End If
                
            Case 8
                If KeyCode = 13 Then
                    ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
                    ZAgrega = Left$(Clientes.Text, 1) + ZSuma
                    ZLong = Len(ZAgrega)
                    If Right$(DBGrid1.Text, ZLong) <> ZAgrega Then
                        DBGrid1.Text = DBGrid1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
                    End If
                    DBGrid1.Col = 9
                    KeyCode = 0
                End If
                
            Case 9
                If KeyCode = 13 Then
                    iRow = DBGrid1.Row
                    DBGrid1.Col = 5
                    XTipo = DBGrid1.Text
                    DBGrid1.Col = 9
                    DBGrid1.Text = Alinea("###,###.##", DBGrid1.Text)
                    Call Suma_Datos
                    DBGrid1.Row = iRow
                    
                    If Val(XTipo) = 4 Then
                        Cuenta1.Text = WCuenta(DBGrid1.Row)
                        IngreCuenta.Visible = True
                        Cuenta1.SetFocus
                    End If
                    
                    ZZCuit = ZClaveCheque(DBGrid1.Row + 1, 6)
                    If Val(XTipo) = 2 And ZZCuit = "" Then
                        Cuit.Text = ""
                        IngresaCuit.Visible = True
                        Cuit.SetFocus
                    End If
                    
                    If DBGrid1.Row < 20 Then
                        DBGrid1.Row = DBGrid1.Row + 1
                        DBGrid1.Col = 5
                        KeyCode = 0
                            Else
                        DBGrid1.Col = 5
                        KeyCode = 0
                    End If
                End If

            Case Else
                
    End Select
                        
    ZZDa = Len(DBGrid1.Text)
    If Len(DBGrid1.Text) = 30 And UCase(Left$(DBGrid1.Text, 1)) = "C" Then
    
        Lectora.Text = "c" + Mid$(DBGrid1.Text, 2, 29) + "e"
        
        ZEntra = "S"
    
        Sql1 = "Select *"
        Sql2 = " FROM Recibos"
        Sql3 = " Where Recibos.ClaveCheque = " + "'" + Lectora.Text + "'"
        spRecibos = Sql1 + Sql2 + Sql3
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            m1$ = "Cheque ya cargado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            ZEntra = "N"
            rstRecibos.Close
        End If
    
        Sql1 = "Select *"
        Sql2 = " FROM RecibosProvi"
        Sql3 = " Where RecibosProvi.ClaveCheque = " + "'" + Lectora.Text + "'"
        spRecibosProvi = Sql1 + Sql2 + Sql3
        Set rstRecibosProvi = db.OpenRecordset(spRecibosProvi, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibosProvi.RecordCount > 0 Then
            m1$ = "Cheque ya cargado"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            ZEntra = "N"
            rstRecibosProvi.Close
        End If
        
        If ZEntra = "S" Then
            For ZZCiclo = 1 To 100
                If ZClaveCheque(ZZCiclo, 1) = Lectora.Text Then
                    m1$ = "Cheque ya cargado"
                    A% = MsgBox(m1$, 0, "Ingreso de Recibos")
                    ZEntra = "N"
                End If
            Next ZZCiclo
        End If
        
        If ZEntra = "S" Then
    
            ZNombreBanco = ZBancos(Val(Mid$(Lectora.Text, 2, 3)))
            ZNroCuenta = Mid$(Lectora.Text, 12, 8)
        
            ZZBanco = Mid$(Lectora.Text, 2, 3)
            ZZSucursal = Mid$(Lectora.Text, 5, 3)
            ZZNroCheque = Mid$(Lectora.Text, 12, 8)
            ZZNroCuenta = Mid$(Lectora.Text, 20, 11)

            DBGrid1.Col = 5
            DBGrid1.Text = "02"
            DBGrid1.Col = 8
            DBGrid1.Text = ZNombreBanco
            ZSuma = Mid$(Str$(Val(Right$(Clientes.Text, 5))), 2, 5)
            DBGrid1.Text = DBGrid1.Text + "/" + Left$(Clientes.Text, 1) + ZSuma
            DBGrid1.Col = 6
            DBGrid1.Text = ZNroCuenta
            DBGrid1.Col = 7
            DBGrid1.Text = ""
            Rem DbGrid1.Col = 6
            KeyCode = 0
        
            ZClaveCheque(DBGrid1.Row + 1, 1) = Lectora.Text
            ZClaveCheque(DBGrid1.Row + 1, 2) = ZZBanco
            ZClaveCheque(DBGrid1.Row + 1, 3) = ZZSucursal
            ZClaveCheque(DBGrid1.Row + 1, 4) = ZZNroCheque
            ZClaveCheque(DBGrid1.Row + 1, 5) = ZZNroCuenta
            ZClaveCheque(DBGrid1.Row + 1, 6) = ""
        
            ZZClave = ZZBanco + ZZSucursal + ZZNroCuenta
            ZZCuit = ""
        
            ZSql = "Select *"
            ZSql = ZSql + " FROM Cuit"
            ZSql = ZSql + " Where Cuit.Clave = " + "'" + ZZClave + "'"
            spCuit = ZSql
            Set rstCuit = db.OpenRecordset(spCuit, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuit.RecordCount > 0 Then
                ZZCuit = Trim(rstCuit!Cuit)
                rstCuit.Close
            End If
            
            ZClaveCheque(DBGrid1.Row + 1, 6) = ZZCuit
            Toto.Text = ""
            Toto.SetFocus
            
                Else
                
            DBGrid1.Col = 5
            DBGrid1.Text = ""
            DBGrid1.Col = 4
            Lectora.Visible = False
            DBGrid1.SetFocus
            
        End If
        
    End If
    
End Sub



