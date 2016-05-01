VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListgas 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Gastos de Importacion por Carpeta"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2055
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   4815
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Carpeta"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Carpeta"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WMovgas.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "listgas.frx":0000
      Left            =   840
      List            =   "listgas.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListgas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Gastos de Importacion por Carpeta"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Movgas.Carpeta} in " + Desde.Text + " to " + Hasta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Movgas.Carpeta, Movgas.Fecha, Movgas.Orden, Movgas.Concepto, Movgas.Importe, Movgas.Proveedor, Movgas.Origen, Movgas.Moneda, Movgas.ImpoDerechos, " _
                    + "Gasimpo.Nombre, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Movgas Movgas, " _
                    + DSQ + ".dbo.Gasimpo Gasimpo, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Movgas.Concepto = Gasimpo.Codigo AND " _
                    + "Movgas.Proveedor = Proveedor.Proveedor AND " _
                    + "Movgas.Carpeta >= 0 AND " _
                    + "Movgas.Carpeta <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgListgas.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

