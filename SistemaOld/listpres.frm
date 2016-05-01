VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListpres 
   Caption         =   "Listado de Prestamos entre Plantas"
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
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox Desti 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Text            =   " "
         Top             =   1200
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
         Left            =   3120
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Hacia"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   975
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
      ReportFileName  =   "WListpres.rpt"
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
      Height          =   3360
      ItemData        =   "listpres.frx":0000
      Left            =   120
      List            =   "listpres.frx":0007
      TabIndex        =   3
      Top             =   2760
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
Attribute VB_Name = "PrgListpres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Desde1 As String
Private Hasta1 As String
Private Producto As String
Private Costo As Double
Private WVector(10000) As String
Private Auxiliar(100, 7) As String
Private WCodigo As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstMovguia As Recordset
Dim spMovguia As String
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
            !varios = "del " + Desde.Text + " al " + Hasta.Text
            .Update
        End If
    End With

    Desde1 = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    Hasta1 = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Select Case Desti.ListIndex
        Case 0
            Desde2 = "0"
            HAsta2 = "9"
        Case 1
            Desde2 = "1"
            HAsta2 = "1"
        Case 2
            Desde2 = "2"
            HAsta2 = "2"
        Case 3
            Desde2 = "3"
            HAsta2 = "3"
        Case 4
            Desde2 = "4"
            HAsta2 = "4"
        Case 5
            Desde2 = "5"
            HAsta2 = "5"
        Case 6
            Desde2 = "6"
            HAsta2 = "6"
        Case 7
            Desde2 = "7"
            HAsta2 = "7"
        Case 8
            Desde2 = "8"
            HAsta2 = "8"
        Case Else
    End Select
    
    Erase WVector
    Renglon = 0

    spMovguia = "ListaMovguiaTotal"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                    If Desde1 <= rstMovguia!FechaOrd And Hasta1 >= rstMovguia!FechaOrd Then
                        If Val(Desde2) <= rstMovguia!Destino And Val(HAsta2) >= rstMovguia!Destino Then
                            If rstMovguia!Tipo = "T" Then
                                If rstMovguia!Codigo > 90000 Then
                                    If rstMovguia!Tipomov = 0 Then
                                        Renglon = Renglon + 1
                                        WVector(Renglon) = rstMovguia!Terminado
                                    End If
                                End If
                            End If
                        End If
                    End If
                    .MoveNext
                            Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    For Da = 1 To Renglon
    
        XEmpresa = WEmpresa
        
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    
        WCodigo = WVector(Da)
        
        Call Calcula_Costo(WCodigo, Costo)
        WCosto = Str$(Costo)
        
        Select Case Val(XEmpresa)
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
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
                        
        XParam = "'" + WCodigo + "','" _
                    + WCosto + "'"
        spTerminado = "ModificaTerminadoCosto " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da
    
    
    Erase WVector
    Renglon = 0

    spMovguia = "ListaMovguiaTotal"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                    If Desde1 <= rstMovguia!FechaOrd And Hasta1 >= rstMovguia!FechaOrd Then
                        If Val(Desde2) <= rstMovguia!Destino And Val(HAsta2) >= rstMovguia!Destino Then
                            If rstMovguia!Tipo = "M" Then
                                If rstMovguia!Codigo > 90000 Then
                                    If rstMovguia!Tipomov = 0 Then
                                        Renglon = Renglon + 1
                                        WVector(Renglon) = rstMovguia!Articulo
                                    End If
                                End If
                            End If
                        End If
                    End If
                    .MoveNext
                            Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    For Da = 1 To Renglon
    
        XEmpresa = WEmpresa
        
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    
        WCodigo = WVector(Da)
        WCosto = ""
        
        spArticulo = "ConsultaArticulo " + "'" + WCodigo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = Str$(rstArticulo!Costo2)
            rstArticulo.Close
        End If
        
        Select Case Val(XEmpresa)
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
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
        End Select
                        
        XParam = "'" + WCodigo + "','" _
                    + WCosto + "'"
        spArticulo = "ModificaarticuloCostoImpre " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Next Da
    
    Listado.WindowTitle = "Listado de Prestamos entre Plantas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Guia.fechaord} in " + Chr$(34) + Desde1 + Chr$(34) + " to " + Chr$(34) + Hasta1 + Chr$(34)
    Dos = " and {Guia.destino} in " + Desde2 + " to " + HAsta2
    Tres = "and {Guia.tipomov} = 0 "
    Cuatro = " and {Guia.Codigo} > 900000"
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Guia.Tipomov, Guia.Codigo, Guia.Fecha, Guia.Tipo, Guia.Articulo, Guia.Terminado, Guia.Cantidad, Guia.Fechaord, Guia.Observaciones, Guia.Destino, " _
                        + "Articulo.Descripcion, Articulo.Costo2, " _
                        + "Terminado.Descripcion, Terminado.Costo " _
                        + "From " _
                        + DSQ + ".dbo.Guia Guia, " _
                        + DSQ + ".dbo.Articulo Articulo, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Guia.Articulo = Articulo.Codigo AND " _
                        + "Guia.Terminado = Terminado.Codigo AND " _
                        + "Guia.Tipomov = 0 AND " _
                        + "Guia.Codigo >= 900000 AND " _
                        + "Guia.Fechaord >= '" + Desde1 + "' AND " _
                        + "Guia.Fechaord <= '" + Hasta1 + "' AND " _
                        + "Guia.Destino >= " + Desde2 + " AND Guia.Destino <= " + HAsta2
    
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListpres.Hide
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
            PrgListpres.Caption = "Listado de Prestamos entre Plantas :  " + !Nombre
        End If
    End With

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Desti.Clear
    
    Desti.AddItem "A Todas"
    Desti.AddItem "Surfactan"
    Desti.AddItem "Pelitall"
    Desti.AddItem "Surfactan II"
    Desti.AddItem "Pelitall II"
    Desti.AddItem "Surfactan III"
    Desti.AddItem "Surfactan IV"
    Desti.AddItem "Surfactan V"
    Desti.AddItem "Pelitall V"
    Desti.AddItem "Pelitall IV"
    Desti.AddItem "Surfactan VI"
    Desti.AddItem "Surfactan VII"
    
    Desti.ListIndex = 0
    
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

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = 1
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        XVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub

