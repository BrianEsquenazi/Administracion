VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImpreCargaIVersion 
   AutoRedraw      =   -1  'True
   Caption         =   "Impresion de Registro de Produccion"
   ClientHeight    =   2775
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2775
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   7
         Text            =   " "
         Top             =   1080
         Width           =   855
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   300
         Left            =   2400
         TabIndex        =   0
         Top             =   600
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
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton AceptaII 
         Caption         =   "Aceptar"
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
         Left            =   4320
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wficter.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImpreCargaIVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String

Dim XParam As String
Dim ZDesdePaso As String
Dim ZHastaPaso As String

Dim ZHumedad(100) As String
Dim ZImpreCarga(200, 6) As String
Dim ZImpreCargaI(100, 20) As String
Dim ZImpreMetodo(100) As String
Dim ZDesTerminado As String
Dim ZFabrica As String
Dim ZZCantidad As Integer

Dim ZVector(100, 10) As String


Private Sub AceptaII_Click()

    Terminado.Text = UCase(Terminado.Text)

    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZDesTerminado = rstTerminado!Descripcion
        ZFabrica = Str$(rstTerminado!fabrica)
        rstTerminado.Close
    End If


    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.WindowTitle = "Impresion de Registro de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    ZDesdePaso = "0"
    ZHastaPaso = "999"
    
    Rem Listado.SQLQuery = "SELECT CargaIIIVersion.Terminado, CargaIIIVersion.Version, CargaIIIVersion.Paso, CargaIIIVersion.Renglon, CargaIIIVersion.Articulo, CargaIIIVersion.PTerminado, CargaIIIVersion.Letra, CargaIIIVersion.Descripcion, CargaIIIVersion.Cantidad, CargaIIIVersion.Item, CargaIIIVersion.DesEpp, CargaIIIVersion.CorteItem, CargaIIIVersion.Terminado " _
    REM     + "From " _
    REM     + DSQ + ".dbo.CargaIIIVersion CargaIIIVersion " _
    REM     + "Where " _
    REM     + "CargaIIIVersion.Terminado >= '" + Terminado.Text + "' AND " _
    REM     + "CargaIIIVersion.Terminado <= '" + Terminado.Text + "' AND " _
    REM     + "CargaIIIVersion.Version >= " + Version.Text + " AND " _
    REM     + "CargaIIIVersion.Version <= " + Version.Text + " AND " _
    REM     + "CargaIIIVersion.Paso >= 0 AND " _
    REM     + "CargaIIIVersion.Paso <= 9999"
    
    Listado.ReportFileName = "WImpreProcedimientoVersion.rpt"
    
    Uno = "{CargaIIIVersion.Paso} in " + ZDesdePaso + " to " + ZHastaPaso
    Dos = " and {CargaIIIVersion.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Tres = " and {CargaIIIversion.Version} in " + Version.Text + " to " + Version.Text
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
    
    Listado.Action = 1

End Sub

Private Sub CANCELA_Click()
    PrgImpreCargaIVersion.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Terminado.Text = "  -     -   "
    Version.Text = ""
    
    Panta.Value = False
    Impresora.Value = True
    
End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Version.SetFocus
    End If
    If KeyAscii = 27 Then
        Terminado.Text = "  -     -   "
    End If
End Sub

Private Sub Version_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.SetFocus
    End If
    If KeyAscii = 27 Then
        Version.Text = ""
    End If
End Sub

