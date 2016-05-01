VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdprv 
   Caption         =   "Listado de  Ordenes de Compra por Proveedor"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Fecha"
      Height          =   2775
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Hastaprov 
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   16
         Text            =   " "
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Desdeprov 
         Height          =   285
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   15
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Proveedor"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wOrdprv.rpt"
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
      ItemData        =   "ordprv.frx":0000
      Left            =   840
      List            =   "ordprv.frx":0007
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
Attribute VB_Name = "PrgOrdprv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Desde1 As String
Private Hasta1 As String

Private Sub Acepta_Click()

    Desde1 = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    Hasta1 = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Listado.WindowTitle = "Listado de Ordenes de Compra por Proveedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Orden.FechaOrd} in " + Chr$(34) + Desde1 + Chr$(34) + " to " + Chr$(34) + Hasta1 + Chr$(34)
    Dos = " and {Orden.Proveedor} in " + Chr$(34) + DesdeProv.Text + Chr$(34) + " to " + Chr$(34) + HastaProv.Text + Chr$(34)
    Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.FechaOrd, Orden.Liberada, Orden.Devuelta, Orden.Fechaentrega, " _
                        + "Articulo.Descripcion, " _
                        + "Proveedor.Nombre " _
                        + "From " _
                        + DSQ + ".dbo.Orden Orden, " _
                        + DSQ + ".dbo.Articulo Articulo, " _
                        + DSQ + ".dbo.Proveedor Proveedor " _
                        + "Where " _
                        + "Orden.Articulo = Articulo.Codigo AND " _
                        + "Orden.Proveedor = Proveedor.Proveedor AND " _
                        + "Orden.Proveedor >= '" + DesdeProv.Text + "' AND " _
                        + "Orden.Proveedor <= '" + HastaProv.Text + "' AND " _
                        + "Orden.FechaOrd >= '" + Desde1 + "' AND " _
                        + "Orden.FechaOrd <= '" + Hasta1 + "' "
                        
    Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgOrdprv.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeProv.SetFocus
    End If
End Sub

Private Sub DesdeProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProv.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdprv.Caption = "Listado de Ordenes de Compra por Proveedor :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeProv.Text = "0"
    HastaProv.Text = "99999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


