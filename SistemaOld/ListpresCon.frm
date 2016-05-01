VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListpresCon 
   Caption         =   "Listado de Prestamos entre Plantas Conslidado"
   ClientHeight    =   6795
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   8235
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6015
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   5
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
      ReportFileName  =   "Listprescon.rpt"
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
      TabIndex        =   3
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
      Height          =   3360
      ItemData        =   "ListpresCon.frx":0000
      Left            =   120
      List            =   "ListpresCon.frx":0007
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   7815
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
Attribute VB_Name = "PrgListpresCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DesdeFecha As String
Private HastaFecha As String
Private Producto As String
Private Costo As Double
Private WVector(10000, 20) As String
Private Auxiliar(100, 7) As String
Private WCodigo As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPresCon As Recordset
Dim spPresCon As String
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim XParam As String

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Varios = "del " + Desde.Text + " al " + Hasta.Text
            .Update
        End If
    End With
    
    spPresCon = "BorrarPresCon "
    Set rstPresCon = db.OpenRecordset(spPresCon, dbOpenSnapshot, dbSQLPassThrough)

    DesdeFecha = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    HastaFecha = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    EmpresaAnterior = WEmpresa
    Erase WVector
    Renglon = 0
    
    For Ciclo = 1 To 8
    
        Select Case Ciclo
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
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
    
        spPrestamo = "ListaPrestamoTotal"
        Set rstPrestamo = db.OpenRecordset(spPrestamo, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrestamo.RecordCount > 0 Then
            With rstPrestamo
                .MoveFirst
                Do
                    If .EOF = False Then
                        If HastaFecha >= rstPrestamo!OrdFecha Then
                            Renglon = Renglon + 1
                            WVector(Renglon, 1) = rstPrestamo!Codigo
                            WVector(Renglon, 2) = rstPrestamo!Fecha
                            WVector(Renglon, 3) = rstPrestamo!Observaciones
                            WVector(Renglon, 4) = rstPrestamo!Tipo
                            WVector(Renglon, 5) = rstPrestamo!articulo
                            WVector(Renglon, 6) = rstPrestamo!Terminado
                            WVector(Renglon, 7) = Str$(rstPrestamo!Cantidad)
                            WVector(Renglon, 8) = Str$(rstPrestamo!Costo)
                            WVector(Renglon, 9) = rstPrestamo!destino
                            WVector(Renglon, 10) = Ciclo
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrestamo.Close
        End If
        
    Next Ciclo
    
    EmpresaNueva = ""
    
    For Ciclo = 1 To Renglon
    
        WCodigo = WVector(Ciclo, 1)
        WFecha = WVector(Ciclo, 2)
        WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WObservaciones = WVector(Ciclo, 3)
        WTipo = WVector(Ciclo, 4)
        WArticulo = WVector(Ciclo, 5)
        WTerminado = WVector(Ciclo, 6)
        WCantidad = WVector(Ciclo, 7)
        WCosto = WVector(Ciclo, 8)
        WDestino = WVector(Ciclo, 9)
        WEmpre = WVector(Ciclo, 10)
        WDescripcion = ""
        Select Case Val(WEmpre)
            Case 1, 3, 5, 6, 7
                WCantidad1 = WCantidad
                WCantidad2 = ""
            Case Else
                WCantidad2 = WCantidad
                WCantidad1 = ""
        End Select
        
        If EmpresaNueva <> WEmpre Then
        
            EmpresaNueva = WEmpre
            
            Select Case Val(EmpresaNueva)
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
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
        End If
            
        If WTipo = "M" Then
        
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WDescripcion = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                    
                Else
                
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WDescripcion = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            
        End If
        
        WVector(Ciclo, 11) = WDescripcion
        
    Next Ciclo
    
    Select Case Val(EmpresaAnterior)
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
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    CostoTotal = 0
    
    For Ciclo = 1 To Renglon
    
        Rem If ciclo = 927 Then Stop
    
        WCodigo = WVector(Ciclo, 1)
        WFecha = WVector(Ciclo, 2)
        WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WObservaciones = WVector(Ciclo, 3)
        WTipo = WVector(Ciclo, 4)
        WArticulo = WVector(Ciclo, 5)
        WTerminado = WVector(Ciclo, 6)
        WCantidad = WVector(Ciclo, 7)
        WCosto = WVector(Ciclo, 8)
        WDestino = WVector(Ciclo, 9)
        WEmpre = WVector(Ciclo, 10)
        WDescripcion = WVector(Ciclo, 11)
        Select Case Val(WEmpre)
            Case 1, 3, 5, 6, 7
                WCantidad1 = WCantidad
                WCantidad2 = ""
            Case Else
                WCantidad2 = WCantidad
                WCantidad1 = ""
        End Select
        
        If DesdeFecha <= WOrdFecha Then
            XParam = "'" + WCodigo + "','" _
                         + WFecha + "','" _
                         + WOrdFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad1 + "','" _
                         + WCantidad2 + "','" _
                         + WCosto + "','" _
                         + WDestino + "','" _
                         + WObservaciones + "','" _
                         + WDescripcion + "'"
                                         
            Set rstPresCon = db.OpenRecordset("AltaPrescon " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            CostoTotal = CostoTotal + (Val(WCantidad1) * Val(WCosto)) + (Val(WCantidad2) * Val(WCosto) * -1)
        End If
        
    Next Ciclo
    
    If CostoTotal <> 0 Then
    
        WCodigo = ""
        WFecha = "00/00/0000"
        WOrdFecha = "0000000"
        WObservaciones = ""
        WTipo = ""
        WArticulo = ""
        WTerminado = ""
        WCantidad = "1"
        WCosto = Str$(Abs(CostoTotal))
        WDestino = ""
        WEmpre = ""
        WDescripcion = ""
        If CostoTotal > 0 Then
                WCantidad1 = "1"
                WCantidad2 = ""
             Else
                WCantidad2 = "1"
                WCantidad1 = ""
        End If
        
        XParam = "'" + WCodigo + "','" _
                         + WFecha + "','" _
                         + WOrdFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad1 + "','" _
                         + WCantidad2 + "','" _
                         + WCosto + "','" _
                         + WDestino + "','" _
                         + WObservaciones + "','" _
                         + WDescripcion + "'"
                                         
        Set rstPresCon = db.OpenRecordset("AltaPrescon " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    
    Listado.WindowTitle = "Listado de Prestamos entre Plantas Consolidado"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT PresCon.Codigo, PresCon.Fecha, PresCon.OrdFecha, PresCon.Tipo, PresCon.Articulo, PresCon.Terminado, PresCon.Cantidad1, PresCon.Cantidad2, PresCon.Costo, PresCon.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.PresCon PresCon " _
                        + "Where " _
                        + "PresCon.Codigo >= 0 AND PresCon.Codigo <= 999999"
    
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
    PrgListpresCon.Hide
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

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListpresCon.Caption = "Listado de Prestamos entre Plantas :  " + !Nombre
        End If
    End With

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
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

