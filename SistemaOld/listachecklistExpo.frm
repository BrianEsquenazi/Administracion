VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaCheckListExpo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Check List de Hojas de Ruta"
   ClientHeight    =   2400
   ClientLeft      =   1875
   ClientTop       =   2520
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         Mask            =   "##/##/####"
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1440
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
         Left            =   3960
         TabIndex        =   6
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
         Left            =   3960
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinfpend.rpt"
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaCheckListExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim Vector(10000, 4) As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstArticulo As Recordset
Dim spArticulo As String

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Listado.WindowTitle = "Listado de Check-List de Hojas de Ruta"
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
    
    Listado.ReportFileName = "ListaCheckListExpo.rpt"
    
    Listado.SQLQuery = "SELECT ChecklistExpo.Hoja, ChecklistExpo.Renglon, ChecklistExpo.Fecha, ChecklistExpo.OrdFecha, ChecklistExpo.Expreso, ChecklistExpo.Chapa, ChecklistExpo.Chofer, ChecklistExpo.Item1, ChecklistExpo.Item2, ChecklistExpo.Item3, ChecklistExpo.Item4, ChecklistExpo.Item5, ChecklistExpo.Item6, ChecklistExpo.Item7, ChecklistExpo.Item8, ChecklistExpo.Placa, ChecklistExpo.Rombo, ChecklistExpo.Observaciones, ChecklistExpo.DesExpreso " _
        + "From " _
        + DSQ + ".dbo.ChecklistExpo ChecklistExpo " _
        + "Where " _
        + "ChecklistExpo.OrdFecha >= '" + WDesde + "' AND " _
        + "ChecklistExpo.OrdFecha <= '" + WHasta + "' AND +" _
        + "ChecklistExpo.Expreso > 0"
        
    Listado.GroupSelectionFormula = "{ChecklistExpo.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
                        
    Listado.Connect = Connect()
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListaCheckListExpo.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

