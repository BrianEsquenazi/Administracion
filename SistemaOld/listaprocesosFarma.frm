VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaProcesosFarma 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Procesos de Fabricacion"
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1440
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
         Left            =   4200
         TabIndex        =   4
         Top             =   480
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
         Left            =   4200
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   960
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
      Begin VB.Label Label2 
         Caption         =   "Hasta"
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
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
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
         Width           =   975
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
Attribute VB_Name = "PrgListaProcesosFarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Listado.WindowTitle = "Listado de Procesos de Fabricacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaIII.Terminado, CargaIII.Paso, CargaIII.Renglon, CargaIII.ControlCambio, CargaIII.Version, CargaIII.FechaVersion, " _
            + "Terminado.Descripcion " _
            + "From " _
            + DSQ + ".dbo.CargaIII CargaIII, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "CargaIII.Terminado = Terminado.Codigo AND " _
            + "CargaIII.Terminado >= '" + Desde.Text + "' AND " _
            + "CargaIII.Terminado <= '" + Hasta.Text + "' AND " _
            + "CargaIII.Paso = 1 AND " + _
            "CargaIII.Renglon = 1"
    
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.Connect = Connect()
    Listado.ReportFileName = "ListaProcesosfarma.rpt"
    Listado.Action = 1
    
End Sub

Private Sub CANCELA_Click()
    PrgListaProcesosFarma.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = "  -     -   "
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  -     -   "
    End If
End Sub


