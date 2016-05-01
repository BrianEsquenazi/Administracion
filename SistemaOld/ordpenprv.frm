VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdPenPrv 
   Caption         =   "Listado de Ordenes Pendientes por Proveedor"
   ClientHeight    =   6150
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   8145
   Begin VB.TextBox Ayuda 
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
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2655
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   1920
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
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdPenPrv.rpt"
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "ordpenprv.frx":0000
      Left            =   240
      List            =   "ordpenprv.frx":0007
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   420
      Left            =   5760
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6960
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgOrdPenPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim XParam As String
Dim XIndice As Integer
Dim Vector(10000, 2) As String

Private Sub Acepta_Click()

    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(Hastafecha.Text, 4)
    WMes = Mid$(Hastafecha.Text, 4, 2)
    WDia = Left$(Hastafecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia
    
    XParam = "'" + "'"

    spOrden = "ModificaOrdenSaldo " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    Erase Vector
    Lugar = 0

    spOrden = "ListaOrdenTotal "
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
    
    With rstOrden
         .MoveFirst
         Do
             If .EOF = False Then
                 WClave = rstOrden!Clave
                 WOrden = rstOrden!Orden
                 WFecha2 = rstOrden!fecha2
                 WSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                 If Val(WSaldo) > 0 Then
                    Entra = "S"
                    For XX = 1 To Lugar
                        If Val(Vector(XX, 1)) = WOrden Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XX
                    
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = WOrden
                        Vector(Lugar, 2) = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
                    End If
                    
                 End If
                 .MoveNext
                     Else
                 Exit Do
             End If
         Loop
    End With
    rstOrden.Close
    
    End If
    
    For XX = 1 To Lugar
        WOrden = Vector(XX, 1)
        WFecha2 = Vector(XX, 2)
        XParam = "'" + WOrden + "','" _
                     + WFecha2 + "'"
    
        spOrden = "ModificaOrdenFecha2 " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Next XX

    Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Orden.Saldo} > 0 And "
    Dos = "{Orden.OrdFecha2} in " + Chr$(34) + WDesdeFecha + Chr$(34) + " to " + Chr$(34) + WHastaFecha + Chr$(34) + " and "
    Tres = "{Orden.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Orden.Orden, Orden.Fecha, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Fecha2, Orden.Saldo, Orden.OrdFecha2, " _
                       + "Proveedor.Nombre, " _
                       + "Articulo.Descripcion " _
                       + "From " + DSQ + ".dbo.Orden Orden, " _
                       + DSQ + ".dbo.Proveedor Proveedor, " _
                       + DSQ + ".dbo.Articulo Articulo " _
                       + "Where Orden.Proveedor = Proveedor.Proveedor AND Orden.Articulo = Articulo.Codigo AND Orden.Proveedor >= '" + Desde.Text + "' AND Orden.Proveedor <= '" + Hasta.Text + "' AND Orden.Saldo > 0. AND Orden.OrdFecha2 >= '" + WDesdeFecha + "' AND Orden.OrdFecha2 <= '" + WHastaFecha + "'"
    
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    Desde.SetFocus
    PrgOrdPenPrv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()


    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesdeFecha = WAno + WMes + WDia
    WAno = Right$(Hastafecha.Text, 4)
    WMes = Mid$(Hastafecha.Text, 4, 2)
    WDia = Left$(Hastafecha.Text, 2)
    WHastaFecha = WAno + WMes + WDia
    
    XParam = "'" + "'"

    spOrden = "ModificaOrdenSaldo " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    Erase Vector
    Lugar = 0

    spOrden = "ListaOrdenTotal "
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
    
    With rstOrden
         .MoveFirst
         Do
             If .EOF = False Then
                 WClave = rstOrden!Clave
                 WOrden = rstOrden!Orden
                 WFecha2 = rstOrden!fecha2
                 WSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                 If Val(WSaldo) > 0 Then
                    Entra = "S"
                    For XX = 1 To Lugar
                        If Val(Vector(XX, 1)) = WOrden Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XX
                    
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = WOrden
                        Vector(Lugar, 2) = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
                    End If
                    
                 End If
                 .MoveNext
                     Else
                 Exit Do
             End If
         Loop
    End With
    rstOrden.Close
    
    End If
    
    For XX = 1 To Lugar
        WOrden = Vector(XX, 1)
        WFecha2 = Vector(XX, 2)
        XParam = "'" + WOrden + "','" _
                     + WFecha2 + "'"
    
        spOrden = "ModificaOrdenFecha2 " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    Next XX

    Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Orden.Saldo} > 0 And "
    Dos = "{Orden.Tipo} = 1"
   
    Listado.GroupSelectionFormula = Uno + Dos
    
    Listado.ReportFileName = "WOrdPenPrvNuevo.rpt"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Orden.Orden, Orden.Proveedor, Orden.Articulo, Orden.Cantidad, Orden.Precio, Orden.Saldo, Orden.Tipo, Orden.Carpeta, " _
            + "Proveedor.Nombre, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Orden Orden, " _
            + DSQ + ".dbo.Proveedor Proveedor, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "Orden.Proveedor = Proveedor.Proveedor AND " _
            + "Orden.Articulo = Articulo.Codigo AND " _
            + "Orden.Saldo > 0 AND " _
            + "Orden.Tipo = 1"
    
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1


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
        Desdefecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefecha.Text, Auxi)
        If Auxi = "S" Then
            Hastafecha.SetFocus
                Else
            Desdefecha.SetFocus
        End If
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hastafecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hastafecha.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdPenPrv.Caption = "Listado de Ordenes Pendientes por Proveedor :  " + !Nombre
        End If
    End With
    Desde.Text = ""
    Hasta.Text = "99999999999"
    Desdefecha.Text = "  /  /    "
    Hastafecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


 Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            spProveedor = "ListaProveedoresOrd"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = RstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = RstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus

End Sub



Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(RstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(RstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    RstProveedor.Close
    
    End If

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Desde.Text = WProveedor
                Hasta.Text = WProveedor
                    Else
                Desde.Text = WProveedor
                Hasta.Text = WProveedor
            End If
            
            Ayuda.Visible = False
            Pantalla.Visible = False
        Case Else
    End Select
    
End Sub

