VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCargaIIILabo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Instrucciones de Produccion de P.T."
   ClientHeight    =   8190
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   11685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.CommandButton AgregaRenglon 
      Caption         =   "Agrega Renglon"
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
      Left            =   7440
      TabIndex        =   33
      Top             =   6960
      Width           =   1095
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
      Left            =   9600
      TabIndex        =   32
      Top             =   6360
      Width           =   1215
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
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   30
      Text            =   " "
      Top             =   480
      Width           =   5760
   End
   Begin VB.CommandButton GrabaII 
      Caption         =   "Revalida"
      Height          =   495
      Left            =   7440
      TabIndex        =   29
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   4200
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   27
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1440
         Width           =   2415
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
         TabIndex        =   28
         Top             =   240
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   4575
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   8070
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso de Procedimiento"
      TabPicture(0)   =   "CargaIIILabo.frx":0000
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
      TabPicture(1)   =   "CargaIIILabo.frx":001C
      Tab(1).ControlCount=   5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WVector2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "WTexto12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "WTexto22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "WTexto32"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "WCombo12"
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
         Top             =   480
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
      Left            =   8760
      MaxLength       =   4
      TabIndex        =   8
      Text            =   " "
      Top             =   120
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
      Left            =   120
      TabIndex        =   5
      Top             =   5640
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
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
      ItemData        =   "CargaIIILabo.frx":0038
      Left            =   120
      List            =   "CargaIIILabo.frx":003F
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   1560
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
   Begin VB.Label lblLabels 
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
      Index           =   3
      Left            =   240
      TabIndex        =   31
      Top             =   480
      Width           =   1935
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   9960
      MouseIcon       =   "CargaIIILabo.frx":004D
      MousePointer    =   99  'Custom
      Picture         =   "CargaIIILabo.frx":0357
      ToolTipText     =   "Impresion "
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Paso"
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
      Left            =   7800
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
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image cmdclose1 
      Height          =   480
      Left            =   10800
      MouseIcon       =   "CargaIIILabo.frx":0B99
      MousePointer    =   99  'Custom
      Picture         =   "CargaIIILabo.frx":0EA3
      ToolTipText     =   "Salida"
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   7440
      MouseIcon       =   "CargaIIILabo.frx":16E5
      MousePointer    =   99  'Custom
      Picture         =   "CargaIIILabo.frx":19EF
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   9120
      MouseIcon       =   "CargaIIILabo.frx":2231
      MousePointer    =   99  'Custom
      Picture         =   "CargaIIILabo.frx":253B
      ToolTipText     =   "Consulta de Datos"
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   8280
      MouseIcon       =   "CargaIIILabo.frx":2D7D
      MousePointer    =   99  'Custom
      Picture         =   "CargaIIILabo.frx":3087
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5760
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
Attribute VB_Name = "PrgCargaIIILabo"
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
Dim rsCargaIII As String
Dim rstCargaV As Recordset
Dim rsCargaV As String
Dim rstEspecifUnifica As Recordset
Dim spEspecifUnifica As String
Dim rstEspecifUnificaVersion As Recordset
Dim spEspecifUnificaVersion As String
Dim rstOperador As Recordset
Dim spOperador As String

Private XIndice As Single
Private Clave As String
Private Auxi As String
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Private WGraba As String
Dim ZVersion As String
Dim CargaEmpresa(12, 2) As String

Dim Ciclo As Integer
Dim XPaso As String
Dim Renglon As Integer
Dim ZEnsayo As String
Dim ZValor As String
Dim ZOperador As String
Dim ZGrabaEspe(1000, 3) As String
Dim ZProceso As Integer


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
Private Sub CaratulaPDF_Click()

    Terminado.Text = UCase(Terminado.Text)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " Partida = " + "'" + "0" + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + ZFabrica + "',"
    ZSql = ZSql + " ImprePaso = Paso "
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()

    Listado.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Ensayo, CargaV.Valor, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, " _
            + "Terminado.Descripcion, Terminado.VersionII, Terminado.FechaVersionII " _
            + "From " _
            + DSQ + ".dbo.CargaV CargaV, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "CargaV.Terminado = Terminado.Codigo AND " _
            + "CargaV.Terminado >= '" + Terminado.Text + "' AND " _
            + "CargaV.Terminado <= '" + Terminado.Text + "' AND " _
            + "CargaV.Paso = 99"
    
    Listado.ReportFileName = "ImpreCalidadNuevoOtro.rpt"
    
    Listado.GroupSelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.Destination = 1
    
    ZZImpreAnterior = Printer.DeviceName
    
    Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "CutePDF Writer" + Chr$(34)
    
    Listado.Action = 1
        
    Rem Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + "HP Color LaserJet 2600n" + Chr$(34)
    Shell "RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n " + Chr$(34) + ZZImpreAnterior + Chr$(34)

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Materia Prima a Utilizar"
     Opcion.AddItem "Productos Terminados a Utilizar"
     Opcion.AddItem "Ensayos"

     Opcion.Visible = True
     
End Sub

Private Sub Lista_Click()

    Terminado.Text = UCase(Terminado.Text)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " Partida = " + "'" + "0" + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + ZFabrica + "',"
    ZSql = ZSql + " ImprePaso = Paso "
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()

    Listado.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Ensayo, CargaV.Valor, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, " _
            + "Terminado.Descripcion, Terminado.VersionII, Terminado.FechaVersionII " _
            + "From " _
            + DSQ + ".dbo.CargaV CargaV, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "CargaV.Terminado = Terminado.Codigo AND " _
            + "CargaV.Terminado >= '" + Terminado.Text + "' AND " _
            + "CargaV.Terminado <= '" + Terminado.Text + "' AND " _
            + "CargaV.Paso = 99"
    
    Listado.ReportFileName = "ImpreCalidadNuevoOtro.rpt"
    
    Listado.GroupSelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    
    Listado.Destination = 1
    Listado.Destination = 0
    Listado.Action = 1
    
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    pantalla.Clear
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
                            pantalla.AddItem IngresaItem
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
                            pantalla.AddItem IngresaItem
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
                            pantalla.AddItem IngresaItem
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
            
        Case Else
    End Select
            
    Ayuda.SetFocus
    pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose1_Click()

    Call Limpia_Click
    PrgCargaIIILabo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    If Trim(ControlCambio.Text) = "" Then
        m$ = "Se debe informar el campo Control de Cambio"
        A% = MsgBox(m$, 0, "Especificaciones de Producto Terminado")
        Exit Sub
    End If
    
    Terminado.Text = UCase(Terminado.Text)

    If WGraba <> "S" Then
    
        ZProceso = 0
        Call Ingresa_clave

               Else

        Sql1 = "DELETE CargaV"
        Sql2 = " Where Terminado = " + "'" + Terminado.Text + "'"
        Sql3 = " and Paso = " + "'" + Paso.Text + "'"
        rsCargaV = Sql1 + Sql2 + Sql3
        Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
    
        If Val(Paso.Text) = 99 Then
            ZDesPaso = "CONTROL FINAL"
            ZCorte = "1"
                Else
            ZDesPaso = "CONTROL PARCIAL PASO: " + Trim(Paso.Text)
            ZCorte = "0"
        End If
        
        Erase ZGrabaEspe
        ZLugar = 0
        ZLugarII = 0

        WRenglon = 0
        For iRow = 1 To 100
    
            WVector2.Row = iRow
            
            WVector2.Col = 1
            ZEnsayo = WVector2.Text
        
            WVector2.Col = 2
            ZDesEnsayo = WVector2.Text
        
            WVector2.Col = 3
            ZValor = WVector2.Text
        
            If Val(ZEnsayo) <> 0 Then
        
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
        
                XPaso = Paso.Text
                Call Ceros(XPaso, 4)
                        
                WClave = Terminado.Text + XPaso + Auxi
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CargaV ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Terminado ,"
                ZSql = ZSql + "Paso ,"
                ZSql = ZSql + "DesPaso ,"
                ZSql = ZSql + "ControlCambio ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Ensayo ,"
                ZSql = ZSql + "DesEnsayo ,"
                ZSql = ZSql + "Valor ,"
                ZSql = ZSql + "Corte )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + Terminado.Text + "',"
                ZSql = ZSql + "'" + Paso.Text + "',"
                ZSql = ZSql + "'" + ZDesPaso + "',"
                ZSql = ZSql + "'" + ControlCambio.Text + "',"
                ZSql = ZSql + "'" + Str$(WRenglon) + "',"
                ZSql = ZSql + "'" + ZEnsayo + "',"
                ZSql = ZSql + "'" + ZDesEnsayo + "',"
                ZSql = ZSql + "'" + ZValor + "',"
                ZSql = ZSql + "'" + ZCorte + "')"
            
                rsCargaV = ZSql
                Set rstCargaV = db.OpenRecordset(rsCargaV, dbOpenSnapshot, dbSQLPassThrough)
                
                ZLugarII = ZLugarII + 1
                If ZLugarII = 1 Then
                    ZLugar = ZLugar + 1
                    ZGrabaEspe(ZLugar, 1) = ZEnsayo
                    ZGrabaEspe(ZLugar, 2) = ZValor
                    ZGrabaEspe(ZLugar, 3) = ""
                        Else
                    If ZGrabaEspe(ZLugar, 1) = ZEnsayo Then
                        ZGrabaEspe(ZLugar, 3) = ZValor
                        ZLugarII = 0
                            Else
                        ZLugarII = 1
                        ZLugar = ZLugar + 1
                        ZGrabaEspe(ZLugar, 1) = ZEnsayo
                        ZGrabaEspe(ZLugar, 2) = ZValor
                        ZGrabaEspe(ZLugar, 3) = ""
                    End If
                End If
        
            End If
            
        Next iRow
        
        
        
        
        
        
        
        
        If Val(Paso.Text) = 99 Then
        
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZSql = ""
            ZSql = ZSql & "Select *"
            ZSql = ZSql & " FROM EspecifUnifica"
            ZSql = ZSql & " Where EspecifUnifica.Producto = " + "'" + Terminado.Text + "'"
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecifUnifica.RecordCount > 0 Then
            
                ZEnsayo1 = Str$(rstEspecifUnifica!Ensayo1)
                ZEnsayo2 = Str$(rstEspecifUnifica!Ensayo2)
                ZEnsayo3 = Str$(rstEspecifUnifica!Ensayo3)
                ZEnsayo4 = Str$(rstEspecifUnifica!Ensayo4)
                ZEnsayo5 = Str$(rstEspecifUnifica!Ensayo5)
                ZEnsayo6 = Str$(rstEspecifUnifica!Ensayo6)
                ZEnsayo7 = Str$(rstEspecifUnifica!Ensayo7)
                ZEnsayo8 = Str$(rstEspecifUnifica!Ensayo8)
                ZEnsayo9 = Str$(rstEspecifUnifica!Ensayo9)
                ZEnsayo10 = Str$(rstEspecifUnifica!Ensayo10)
                ZValor1 = rstEspecifUnifica!Valor1
                ZValor2 = rstEspecifUnifica!valor2
                ZValor3 = rstEspecifUnifica!Valor3
                ZValor4 = rstEspecifUnifica!valor4
                ZValor5 = rstEspecifUnifica!valor5
                ZValor6 = rstEspecifUnifica!valor6
                ZValor7 = rstEspecifUnifica!valor7
                ZValor8 = rstEspecifUnifica!valor8
                ZValor9 = rstEspecifUnifica!valor9
                ZValor10 = rstEspecifUnifica!valor10
                ZValor11 = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                ZValor22 = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                ZValor33 = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                ZValor44 = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                ZValor55 = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                ZValor66 = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                ZValor77 = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                ZValor88 = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                ZValor99 = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                ZValor1010 = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                ZVersion = rstEspecifUnifica!Version
                ZFechaInicio = rstEspecifUnifica!Fecha
                ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
            
                rstEspecifUnifica.Close
            
                Call Ceros(ZVersion, 4)
                ZClave = ZVersion + Terminado.Text
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnificaVersion ("
                ZSql = ZSql & "Clave, "
                ZSql = ZSql & "Version, "
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "FechaInicio , "
                ZSql = ZSql & "FechaFinal , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + ZClave + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + Terminado.Text + "',"
                ZSql = ZSql & "'" + ZEnsayo1 + "'," + "'" + ZValor1 + "',"
                ZSql = ZSql & "'" + ZEnsayo2 + "'," + "'" + ZValor2 + "',"
                ZSql = ZSql & "'" + ZEnsayo3 + "'," + "'" + ZValor3 + "',"
                ZSql = ZSql & "'" + ZEnsayo4 + "'," + "'" + ZValor4 + "',"
                ZSql = ZSql & "'" + ZEnsayo5 + "'," + "'" + ZValor5 + "',"
                ZSql = ZSql & "'" + ZEnsayo6 + "'," + "'" + ZValor6 + "',"
                ZSql = ZSql & "'" + ZEnsayo7 + "'," + "'" + ZValor7 + "',"
                ZSql = ZSql & "'" + ZEnsayo8 + "'," + "'" + ZValor8 + "',"
                ZSql = ZSql & "'" + ZEnsayo9 + "'," + "'" + ZValor9 + "',"
                ZSql = ZSql & "'" + ZEnsayo10 + "'," + "'" + ZValor10 + "',"
                ZSql = ZSql & "'" + ZValor11 + "',"
                ZSql = ZSql & "'" + ZValor22 + "',"
                ZSql = ZSql & "'" + ZValor33 + "',"
                ZSql = ZSql & "'" + ZValor44 + "',"
                ZSql = ZSql & "'" + ZValor55 + "',"
                ZSql = ZSql & "'" + ZValor66 + "',"
                ZSql = ZSql & "'" + ZValor77 + "',"
                ZSql = ZSql & "'" + ZValor88 + "',"
                ZSql = ZSql & "'" + ZValor99 + "',"
                ZSql = ZSql & "'" + ZValor1010 + "',"
                ZSql = ZSql & "'" + ZFechaInicio + "',"
                ZSql = ZSql & "'" + ZFechaFinal + "',"
                ZSql = ZSql & "'" + ControlCambio.Text + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
          
                spEspecifUnificaVersion = ZSql
                Set rstEspecifUnificaVersion = db.OpenRecordset(spEspecifUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                
                
                ZVersion = Str$(Val(ZVersion) + 1)
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
                ZEstado = "S"
                WDate = Date$
            
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Producto = " + "'" + Terminado.Text + "',"
                ZSql = ZSql & "Ensayo1 = " + "'" + ZGrabaEspe(1, 1) + "',"
                ZSql = ZSql & "Valor1 = " + "'" + ZGrabaEspe(1, 2) + "',"
                ZSql = ZSql & "Ensayo2 = " + "'" + ZGrabaEspe(2, 1) + "',"
                ZSql = ZSql & "Valor2 = " + "'" + ZGrabaEspe(2, 2) + "',"
                ZSql = ZSql & "Ensayo3 = " + "'" + ZGrabaEspe(3, 1) + "',"
                ZSql = ZSql & "Valor3 = " + "'" + ZGrabaEspe(3, 2) + "',"
                ZSql = ZSql & "Ensayo4 = " + "'" + ZGrabaEspe(4, 1) + "',"
                ZSql = ZSql & "Valor4 = " + "'" + ZGrabaEspe(4, 2) + "',"
                ZSql = ZSql & "Ensayo5 = " + "'" + ZGrabaEspe(5, 1) + "',"
                ZSql = ZSql & "Valor5 = " + "'" + ZGrabaEspe(5, 2) + "',"
                ZSql = ZSql & "Ensayo6 = " + "'" + ZGrabaEspe(6, 1) + "',"
                ZSql = ZSql & "Valor6 = " + "'" + ZGrabaEspe(6, 2) + "',"
                ZSql = ZSql & "Ensayo7 = " + "'" + ZGrabaEspe(7, 1) + "',"
                ZSql = ZSql & "Valor7 = " + "'" + ZGrabaEspe(7, 2) + "',"
                ZSql = ZSql & "Ensayo8 = " + "'" + ZGrabaEspe(8, 1) + "',"
                ZSql = ZSql & "Valor8 = " + "'" + ZGrabaEspe(8, 2) + "',"
                ZSql = ZSql & "Ensayo9 = " + "'" + ZGrabaEspe(9, 1) + "',"
                ZSql = ZSql & "Valor9 = " + "'" + ZGrabaEspe(9, 2) + "',"
                ZSql = ZSql & "Ensayo10 = " + "'" + ZGrabaEspe(10, 1) + "',"
                ZSql = ZSql & "Valor10 = " + "'" + ZGrabaEspe(10, 2) + "',"
                ZSql = ZSql & "WDate = " + "'" + WDate + "',"
                ZSql = ZSql & "Valor11 = " + "'" + ZGrabaEspe(1, 3) + "',"
                ZSql = ZSql & "Valor22 = " + "'" + ZGrabaEspe(2, 3) + "',"
                ZSql = ZSql & "Valor33 = " + "'" + ZGrabaEspe(3, 3) + "',"
                ZSql = ZSql & "Valor44 = " + "'" + ZGrabaEspe(4, 3) + "',"
                ZSql = ZSql & "Valor55 = " + "'" + ZGrabaEspe(5, 3) + "',"
                ZSql = ZSql & "Valor66 = " + "'" + ZGrabaEspe(6, 3) + "',"
                ZSql = ZSql & "Valor77 = " + "'" + ZGrabaEspe(7, 3) + "',"
                ZSql = ZSql & "Valor88 = " + "'" + ZGrabaEspe(8, 3) + "',"
                ZSql = ZSql & "Valor99 = " + "'" + ZGrabaEspe(9, 3) + "',"
                ZSql = ZSql & "Valor1010 = " + "'" + ZGrabaEspe(10, 3) + "',"
                ZSql = ZSql & "Version = " + "'" + ZVersion + "',"
                ZSql = ZSql & "Fecha = " + "'" + ZFecha + "',"
                ZSql = ZSql & "Estado = " + "'" + ZEstado + "',"
                ZSql = ZSql & "ControlCambio = " + "'" + ControlCambio.Text + "',"
                ZSql = ZSql & "Observaciones = " + "'" + ZObservaciones + "'"
                ZSql = ZSql & " Where Producto = " + "'" + Terminado.Text + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                
                        Else
                        
                    
                ZVersion = "1"
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZObservaciones = ""
                ZEstado = "S"
                WDate = Date$
            
                ZSql = ""
                ZSql = ZSql & "INSERT INTO EspecifUnifica ("
                ZSql = ZSql & "Producto, "
                ZSql = ZSql & "Ensayo1, Valor1, "
                ZSql = ZSql & "Ensayo2, Valor2, "
                ZSql = ZSql & "Ensayo3, Valor3, "
                ZSql = ZSql & "Ensayo4, Valor4, "
                ZSql = ZSql & "Ensayo5, Valor5, "
                ZSql = ZSql & "Ensayo6, Valor6, "
                ZSql = ZSql & "Ensayo7, Valor7, "
                ZSql = ZSql & "Ensayo8, Valor8, "
                ZSql = ZSql & "Ensayo9, Valor9, "
                ZSql = ZSql & "Ensayo10, Valor10, "
                ZSql = ZSql & "WDate, "
                ZSql = ZSql & "Valor11 , "
                ZSql = ZSql & "Valor22 , "
                ZSql = ZSql & "Valor33 , "
                ZSql = ZSql & "Valor44 , "
                ZSql = ZSql & "Valor55 , "
                ZSql = ZSql & "Valor66 , "
                ZSql = ZSql & "Valor77 , "
                ZSql = ZSql & "Valor88 , "
                ZSql = ZSql & "Valor99 , "
                ZSql = ZSql & "Valor1010 , "
                ZSql = ZSql & "Version , "
                ZSql = ZSql & "Fecha , "
                ZSql = ZSql & "Estado , "
                ZSql = ZSql & "ControlCambio , "
                ZSql = ZSql & "Observaciones) "
                ZSql = ZSql & "Values ("
                ZSql = ZSql & "'" + Terminado.Text + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(1, 1) + "'," + "'" + ZGrabaEspe(1, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(2, 1) + "'," + "'" + ZGrabaEspe(2, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(3, 1) + "'," + "'" + ZGrabaEspe(3, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(4, 1) + "'," + "'" + ZGrabaEspe(4, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(5, 1) + "'," + "'" + ZGrabaEspe(5, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(6, 1) + "'," + "'" + ZGrabaEspe(6, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(7, 1) + "'," + "'" + ZGrabaEspe(7, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(8, 1) + "'," + "'" + ZGrabaEspe(8, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(9, 1) + "'," + "'" + ZGrabaEspe(9, 2) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(10, 1) + "'," + "'" + ZGrabaEspe(10, 2) + "',"
                ZSql = ZSql & "'" + WDate + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(1, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(2, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(3, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(4, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(5, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(6, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(7, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(8, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(9, 3) + "',"
                ZSql = ZSql & "'" + ZGrabaEspe(10, 3) + "',"
                ZSql = ZSql & "'" + ZVersion + "',"
                ZSql = ZSql & "'" + ZFecha + "',"
                ZSql = ZSql & "'" + ZEstado + "',"
                ZSql = ZSql & "'" + ControlCambio.Text + "',"
                ZSql = ZSql & "'" + ZObservaciones + "')"
           
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE EspecifUnifica SET "
            ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
            ZSql = ZSql + " Where Producto = " + "'" + Terminado.Text + "'"
                            
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
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
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                WCodigo = WProducto
                WVersionII = ZVersion
                WFechaVersionII = ZFecha
                WEstadoII = ZEstado
                WObservaII = ZObservaciones
            
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "VersionII = " + "'" + ZVersion + "',"
                ZSql = ZSql & "FechaVersionII = " + "'" + ZFecha + "',"
                ZSql = ZSql & "EstadoII = " + "'" + ZEstado + "',"
                ZSql = ZSql & "ObservaII = " + "'" + ZObservaciones + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + Terminado.Text + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
            Next Cicla
            
            
        
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        End If

        ZZPasaTerminado = Terminado.Text
        ZZPasaPaso = Paso.Text
        PrgCargaIIIProduccionAutomatico.Show
        PrgCargaIIIProduccionAutomatico.Hide
        Unload PrgCargaIIIProduccionAutomatico
    
        Call Limpia_Click

        WVector1.TopRow = 1
        WVector1.Col = 1
        WVector1.Row = 1
    
        WVector2.TopRow = 1
        WVector2.Col = 1
        WVector2.Row = 1

        Tablas.Tab = 1
        
        Terminado.SetFocus
        
    End If
        
End Sub



Private Sub GrabaII_Click()

    If WGraba <> "S" Then
    
        ZProceso = 1
        Call Ingresa_clave

               Else
        
        If Val(Paso.Text) = 99 Then
        
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            ZEstado = "S"
            
            ZSql = ""
            ZSql = ZSql & "UPDATE EspecifUnifica SET "
            ZSql = ZSql & "Estado = " + "'" + ZEstado + "'"
            ZSql = ZSql & " Where Producto = " + "'" + Terminado.Text + "'"
                        
            spEspecifUnifica = ZSql
            Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
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
            
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                ZSql = ""
                ZSql = ZSql & "UPDATE Terminado SET "
                ZSql = ZSql & "EstadoII = " + "'" + ZEstado + "'"
                ZSql = ZSql & " Where Codigo = " + "'" + Terminado.Text + "'"
                    
                spTerminado = ZSql
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
            Next Cicla
        
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
        End If
    
        Call Limpia_Click

        WVector1.TopRow = 1
        WVector1.Col = 1
        WVector1.Row = 1
    
        WVector2.TopRow = 1
        WVector2.Col = 1
        WVector2.Row = 1

        Tablas.Tab = 1
        
        Terminado.SetFocus
        
    End If
        
End Sub


Private Sub Limpia_Click()
    
    Call Limpia_Vector
    Call Limpia_VectorII
    
    Tablas.Tab = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    WGraba = ""
    
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

Private Sub pantalla_Click()
    pantalla.Visible = False
    Opcion.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Terminado.Text = WIndice.List(Indice)
            Call Terminado_KeyPress(13)
            
        Case 3
            Indice = pantalla.ListIndex
            WEnsayo = WIndice.List(Indice)
            
            WTexto12.Visible = False
            WTexto22.Visible = False
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WEnsayo + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 1
                WVector2.Text = Trim(rstEnsayos!Codigo)
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                WVector2.Col = 3
                rstEnsayos.Close
                Call StartEditII
            End If
            Rem Ayuda.Visible = False
            
        Case Else
    End Select
    Ayuda.Visible = False
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1

    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Paso.Text = ""
    WGraba = ""
    
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
    ZSql = ZSql + " Order by CargaIII.Clave"
    
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
                    
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstCargaIII!Articulo)
            
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstCargaIII!PTerminado)
            
                    WVector1.Col = 3
                    WVector1.Text = Trim(rstCargaIII!Letra)
            
                    WVector1.Col = 4
                    WVector1.Text = Trim(rstCargaIII!Descripcion)
            
                    WVector1.Col = 5
                    WVector1.Text = rstCargaIII!Cantidad
                    WVector1.Text = Pusing("###.####", WVector1.Text)
                    
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
                
                    ControlCambio.Text = IIf(IsNull(rstCargaV!ControlCambio), "", rstCargaV!ControlCambio)
                
                    WVector2.Col = 1
                    WVector2.Text = Trim(Str$(rstCargaV!Ensayo))
            
                    WVector2.Col = 2
                    WVector2.Text = ""
            
                    WVector2.Col = 3
                    WVector2.Text = Trim(rstCargaV!Valor)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaV.Close
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
    
    Tablas.Tab = 1
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
    
    WVector2.TopRow = 1
    WVector2.Col = 1
    WVector2.Row = 1
    
    Call StartEditII
    
    Graba.Enabled = True

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
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
        Terminado.Text = ""
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
            Tablas.Tab = 1
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            WVector2.TopRow = 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEditII
        End If
        
    End If
    If KeyAscii = 27 Then
        Paso.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    pantalla.Clear
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
                            Da = Len(rstTerminado!Descripcion) - WEspacios
                            For aa = 1 To Da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, aa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    pantalla.AddItem IngresaItem
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
                            pantalla.AddItem IngresaItem
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
                            pantalla.AddItem IngresaItem
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
    WVector1.Rows = 160
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
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###.####"
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














Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1
        Case 113
            WTexto12.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 123
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Col > 1 Then
                WVector2.Col = WVector2.Col - 1
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1
        Case 113
            WTexto22.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub Wtexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto32.Text = WVector2.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row + 1
                End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Call Control_CampoII
                If WControlII = "S" Then
                    WVector2.Row = WVector2.Row - 1
                End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_CampoII
                Rem If WControlII = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub Wtexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 3
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControlII = "S"
    Select Case XColumna
        Case 1
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
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ensayos"
            ZSql = ZSql + " Where Ensayos.Codigo = " + "'" + WVector2.Text + "'"
            spEnsayos = ZSql
            Set rstEnsayos = db.OpenRecordset(spEnsayos, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayos.RecordCount > 0 Then
                WVector2.Col = 2
                WVector2.Text = Trim(rstEnsayos!Descripcion)
                rstEnsayos.Close
                    Else
                WControlII = "N"
            End If
    
            Call Conecta_Empresa
            
        Case 3, 6, 7
            Rem If Val(WVector2.Text) <> 0 Then
            Rem     ZCodigo = Val(WVector2.Text)
            Rem     Call Ceros(ZCodigo, 4)
            Rem
            Rem     Sql1 = "Select *"
            Rem     Sql2 = " FROM EquipoFabrica"
            Rem     Sql3 = " Where EquipoFabrica.Codigo = " + "'" + ZCodigo + "'"
            Rem     spEquipoFabrica = Sql1 + Sql2 + Sql3
            Rem     Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstEquipoFabrica.RecordCount > 0 Then
            Rem         rstEquipoFabrica.Close
            Rem     End If
            Rem End If
            
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
    
    RenglonAuxiliar = WVector2.Row

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    HastaRenglon = 0
    For iRow = 100 To 1 Step -1
        
        Ensayo = WVector2.TextMatrix(iRow, 1)
            
        If Ensayo <> "" Then
            HastaRenglon = iRow
            Exit For
        End If
            
    Next iRow
    
    For Ciclo = 1 To HastaRenglon
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        If Ciclo <> RenglonAuxiliar Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 0 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For Da = 0 To WVector2.Cols - 1
            WVector2.Col = Da
            WVector2.Text = WBorraII(Ciclo, Da)
        Next Da
    Next Ciclo
    
    End If
    
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

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 0
            WVector1.Col = 1
            WVector1.Row = 1
        Case 1
            WVector2.Col = 1
            WVector2.Row = 1
            Call StartEditII
        Case Else
    End Select
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
        ZGRABAII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            If ZProceso = 1 Then
                Call GrabaII_Click
                    Else
                Call Graba_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Productos")
            WClave.SetFocus
        End If
        
    End If
End Sub


Private Sub AgregaRenglon_Click()
    Hasta = WVector2.Row

    For iRow = 100 To Hasta Step -1
        WVector2.TextMatrix(iRow, 0) = WVector2.TextMatrix(iRow - 1, 0)
        WVector2.TextMatrix(iRow, 1) = WVector2.TextMatrix(iRow - 1, 1)
        WVector2.TextMatrix(iRow, 2) = WVector2.TextMatrix(iRow - 1, 2)
        WVector2.TextMatrix(iRow, 3) = WVector2.TextMatrix(iRow - 1, 3)
    Next iRow

    WVector2.TextMatrix(Hasta, 0) = ""
    WVector2.TextMatrix(Hasta, 1) = ""
    WVector2.TextMatrix(Hasta, 2) = ""
    WVector2.TextMatrix(Hasta, 3) = ""
    
    WTexto12.Text = ""
    WTexto22.Text = ""

End Sub


