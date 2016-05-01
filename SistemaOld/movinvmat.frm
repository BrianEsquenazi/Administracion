VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovInvMat 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Recuento de Inventario de Materia Prima"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
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
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wMovInvMat.rpt"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "movinvmat.frx":0000
      Left            =   120
      List            =   "movinvmat.frx":0007
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgMovInvMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    Listado.WindowTitle = "Listado de Movimiento de Recuento de Inventario de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Inventario.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Inventario.Numero, Inventario.Tipo, Inventario.Articulo, Inventario.Cantidad, Inventario.Lote, Inventario.Talon, Inventario.Observaciones, Inventario.Ubicacion, " _
                        + "Articulo.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.Inventario Inventario, " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where Inventario.Articulo = Articulo.Codigo AND " _
                        + "Inventario.Tipo = 'M' AND Inventario.Articulo >= 'AA-000--00' AND Inventario.Articulo <= 'ZZ-999-99'"
    
    Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgMovInvMat.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovInvMat.Caption = "Listado de Recuento de Inventario de Materia Prima :  " + !Nombre
        End If
    End With
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstArticulo!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
            Desde.Text = rstArticulo!Codigo
            Hasta.Text = rstArticulo!Codigo
                Else
            Desde.Text = WArticulo
            Hasta.Text = WArticulo
    End If
    Desde.SetFocus
    
    
End Sub

