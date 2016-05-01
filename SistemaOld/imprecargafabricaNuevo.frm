VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImpreCargaFabricaNuevo 
   AutoRedraw      =   -1  'True
   Caption         =   "Impresion de Registro de Produccion"
   ClientHeight    =   2820
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   2820
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1560
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1560
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
      Begin VB.CommandButton Acepta 
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
Attribute VB_Name = "PrgImpreCargaFabricaNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String

Private Sub Acepta_Click()

    Terminado.Text = UCase(Terminado.Text)
    
    Listado.WindowTitle = "Instrucciones de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CargaIV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaIV.Clave, CargaIV.Terminado, CargaIV.Fecha, CargaIV.Version, CargaIV.Etapa, CargaIV.DesTerminado, CargaIV.Impre11, CargaIV.Impre12, CargaIV.Impre13, CargaIV.Impre14, CargaIV.Impre15, CargaIV.Impre16, CargaIV.Impre17, CargaIV.Impre18, CargaIV.Impre19, CargaIV.Impre2, CargaIV.Impre31, CargaIV.Impre32, CargaIV.Impre33, CargaIV.Impre34, CargaIV.Impre35, CargaIV.Impre36, CargaIV.Impre37, CargaIV.Impre38, CargaIV.Impre39, CargaIV.Impre41, CargaIV.Impre42, CargaIV.Impre43, CargaIV.Impre44, CargaIV.Impre45, CargaIV.Impre46, CargaIV.Impre47, CargaIV.Impre48, CargaIV.Impre49, CargaIV.Impre51, CargaIV.Impre52, CargaIV.Impre53, CargaIV.Impre54, CargaIV.Impre55, CargaIV.Impre56, CargaIV.Impre57, CargaIV.Impre58, CargaIV.Impre59, CargaIV.Impre6 " _
                    + "From " _
                    + DSQ + ".dbo.CargaIV CargaIV " _
                    + "Where " _
                    + "CargaIV.Terminado >= '" + Terminado.Text + "' AND " _
                    + "CargaIV.Terminado <= '" + Terminado.Text + "'"

    Listado.Connect = Connect()
    Listado.ReportFileName = "ImpreProceso.rpt"
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    Terminado.SetFocus
    PrgImpreCargaFabrica.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Panta.Value = False
    Impresora.Value = True
End Sub

