VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovvar2 
   Caption         =   "Listado de  Movimientos Varios de Producto Terminado"
   ClientHeight    =   6795
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   8235
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3015
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin MSMask.MaskEdBox HastaArt 
         Height          =   300
         Left            =   2280
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeArt 
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   " "
         Top             =   2280
         Width           =   2175
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
         Left            =   3960
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Producto Terminado"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Producto Termnado"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
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
      Left            =   6960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wmovvar2.rpt"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   720
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
      ItemData        =   "movvar2.frx":0000
      Left            =   120
      List            =   "movvar2.frx":0007
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6960
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgMovvar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Desde1 As String
Private Hasta1 As String
Dim rstTerminado As Recordset
Dim spTerminado As String

Private Sub Acepta_Click()

    DesdeArt.Text = UCase(DesdeArt.Text)
    HastaArt.Text = UCase(HastaArt.Text)

    Desde1 = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    Hasta1 = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Listado.WindowTitle = "Listado de Movimientos Varios de Producto Terminado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Lugar = Tipo.ListIndex
    
    Select Case Lugar
        Case 0
            Uno = "{movvar.movi} in " + Chr$(34) + "A" + Chr$(34) + " to " + Chr$(34) + "Z" + Chr$(34)
        Case 1
            Uno = "{movvar.movi} in " + Chr$(34) + "E" + Chr$(34) + " to " + Chr$(34) + "E" + Chr$(34)
        Case 2
            Uno = "{movvar.movi} in " + Chr$(34) + "S" + Chr$(34) + " to " + Chr$(34) + "S" + Chr$(34)
        Case Else
    End Select
        
    Dos = " and {movvar.tipo} = " + Chr$(34) + "T" + Chr$(34)
    Tres = " and {movvar.fechaord} in " + Chr$(34) + Desde1 + Chr$(34) + " to " + Chr$(34) + Hasta1 + Chr$(34)
    Cuatro = " and {movvar.terminado} in " + Chr$(34) + DesdeArt + Chr$(34) + " to " + Chr$(34) + HastaArt + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Movvar.Codigo, Movvar.Fecha, Movvar.Tipo, Movvar.Articulo, Movvar.Terminado, Movvar.Cantidad, Movvar.FechaOrd, Movvar.Movi, Terminado.Descripcion " _
                        + "From " + DSQ + ".dbo.Movvar Movvar, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where Movvar.Terminado = Terminado.Codigo AND Movvar.Tipo = 'T' AND Movvar.Articulo >= ' ' AND Movvar.Articulo <= 'ZZ-ZZZ-ZZZ' AND Movvar.FechaOrd >= '00000000' AND Movvar.FechaOrd <= '99999999' AND Movvar.Movi >= 'A' AND Movvar.Movi <= 'Z'"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgMovvar2.Hide
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
        DesdeArt.SetFocus
    End If
End Sub

Private Sub DesdeArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArt.Text = UCase(DesdeArt.Text)
        HastaArt.SetFocus
    End If
End Sub

Private Sub hastaart_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArt.Text = UCase(HastaArt.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgMovvar2.Caption = "Listado de Movimientos Varios de Producto Terminado :  " + !Nombre
        End If
    End With

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeArt.Text = "  -     -   "
    HastaArt.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Total"
    Tipo.AddItem "Entradas"
    Tipo.AddItem "Salidas"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstTerminado!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesdeArt.Text = rstTerminado!Codigo
        HastaArt.Text = rstTerminado!Codigo
            Else
        DesdeArt.Text = Claveven$
        HastaArt.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


