VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdartCon 
   Caption         =   "Listado de  Ordenes de Compra (Consolidado)"
   ClientHeight    =   4500
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4500
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox TipoOrden 
         Height          =   315
         Left            =   3960
         TabIndex        =   19
         Top             =   3480
         Width           =   2535
      End
      Begin VB.ComboBox Planta 
         Height          =   315
         Left            =   3960
         TabIndex        =   18
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   3960
         TabIndex        =   17
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Desdeprov 
         Height          =   285
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   14
         Text            =   " "
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Hastaprov 
         Height          =   285
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   13
         Text            =   " "
         Top             =   3000
         Width           =   1455
      End
      Begin MSMask.MaskEdBox HastaArt 
         Height          =   300
         Left            =   2040
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DesdeArt 
         Height          =   300
         Left            =   2040
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   4680
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Desde Proveedor"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta Proveedor"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Articulo"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Articulo"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdartcon.rpt"
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
Attribute VB_Name = "PrgOrdartCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Desde1 As String
Private Hasta1 As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim Empe(12, 2) As String

Private Sub Acepta_Click()

    DesdeArt.Text = UCase(DesdeArt.Text)
    HastaArt.Text = UCase(HastaArt.Text)

    Desde1 = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    Hasta1 = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    With rstWOrden
        .Index = "Orden"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    XEmpresa = WEmpresa
        
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    If Planta.ListIndex = 1 Then
        Select Case Val(WEmpresa)
            Case 1
                Empe(1, 1) = "0001"
                Empe(1, 2) = "Empresa01"
                XHasta = 1
            Case 2
                Empe(1, 1) = "0002"
                Empe(1, 2) = "Empresa02"
                XHasta = 1
            Case 3
                Empe(1, 1) = "0003"
                Empe(1, 2) = "Empresa03"
                XHasta = 1
            Case 4
                Empe(1, 1) = "0004"
                Empe(1, 2) = "Empresa04"
                XHasta = 1
            Case 5
                Empe(1, 1) = "0005"
                Empe(1, 2) = "Empresa05"
                XHasta = 1
            Case 6
                Empe(1, 1) = "0006"
                Empe(1, 2) = "Empresa06"
                XHasta = 1
            Case 7
                Empe(1, 1) = "0007"
                Empe(1, 2) = "Empresa07"
                XHasta = 1
            Case 8
                Empe(1, 1) = "0008"
                Empe(1, 2) = "Empresa08"
                XHasta = 1
            Case 9
                Empe(1, 1) = "0009"
                Empe(1, 2) = "Empresa09"
                XHasta = 1
            Case 10
                Empe(1, 1) = "0010"
                Empe(1, 2) = "Empresa10"
                XHasta = 1
            Case 11
                Empe(1, 1) = "0011"
                Empe(1, 2) = "Empresa11"
                XHasta = 1
            Case Else
        End Select
    End If

    If Tipo.ListIndex = 0 Then

        For a = 1 To XHasta
    
            WEmpresa = Empe(a, 1)
            txtOdbc = Empe(a, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spOrden = "ListaOrdenTotal "
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
    
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
    
                            If DesdeArt.Text <= rstOrden!Articulo And HastaArt.Text >= rstOrden!Articulo Then
                                If Desde1 <= rstOrden!FechaOrd And Hasta1 >= rstOrden!FechaOrd Then
                                    If TipoOrden.ListIndex = 4 Or TipoOrden.ListIndex = rstOrden!Tipo Then
                                        WOrden = rstOrden!Orden
                                        WArticulo = rstOrden!Articulo
                                        WProveedor = rstOrden!Proveedor
                                        WFecha = rstOrden!Fecha
                                        WCantidad = rstOrden!Cantidad
                                        WPrecio = rstOrden!Precio
                                        WLiberada = rstOrden!Liberada
                                        WDevuelta = rstOrden!devuelta
                                        WFechaEntrega = rstOrden!FechaEntrega
                                        WDesArticulo = ""
                                        WDEsProveedor = ""
                                
                                        With rstWOrden
                                            .AddNew
                                            !Orden = WOrden
                                            !Articulo = WArticulo
                                            !Proveedor = WProveedor
                                            !Fecha = WFecha
                                            !Cantidad = WCantidad
                                            !Precio = WPrecio
                                            !Liberada = WLiberada
                                            !devuelta = WDevuelta
                                            !FechaEntrega = WFechaEntrega
                                            !DesArticulo = ""
                                            !DesProveedor = ""
                                            .Update
                                        End With
                                    End If
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
            
        Next a
        
            Else
        
        For a = 1 To XHasta
    
            WEmpresa = Empe(a, 1)
            txtOdbc = Empe(a, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
            spOrden = "ListaOrdenTotal "
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
    
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
    
                            If DesdeProv.Text <= rstOrden!Proveedor And HastaProv.Text >= rstOrden!Proveedor Then
                                If Desde1 <= rstOrden!FechaOrd And Hasta1 >= rstOrden!FechaOrd Then
                                    If TipoOrden.ListIndex = 4 Or TipoOrden.ListIndex = rstOrden!Tipo Then
                                        WOrden = rstOrden!Orden
                                        WArticulo = rstOrden!Articulo
                                        WProveedor = rstOrden!Proveedor
                                        WFecha = rstOrden!Fecha
                                        WCantidad = rstOrden!Cantidad
                                        WPrecio = rstOrden!Precio
                                        WLiberada = rstOrden!Liberada
                                        WDevuelta = rstOrden!devuelta
                                        WFechaEntrega = rstOrden!FechaEntrega
                                        WDesArticulo = ""
                                        WDEsProveedor = ""
                                
                                        With rstWOrden
                                            .AddNew
                                            !Orden = WOrden
                                            !Articulo = WArticulo
                                            !Proveedor = WProveedor
                                            !Fecha = WFecha
                                            !Cantidad = WCantidad
                                            !Precio = WPrecio
                                            !Liberada = WLiberada
                                            !devuelta = WDevuelta
                                            !FechaEntrega = WFechaEntrega
                                            !DesArticulo = ""
                                            !DesProveedor = ""
                                            .Update
                                        End With
                                    End If
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
            
        Next a
        
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
    
    With rstWOrden
        .Index = "Orden"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WProveedor = !Proveedor
                WArticulo = !Articulo
                
                WDEsProveedor = ""
                spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WDEsProveedor = RstProveedor!Nombre
                    RstProveedor.Close
                End If
                
                WDesArticulo = ""
                spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDesArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                !DesArticulo = WDesArticulo
                !DesProveedor = WDEsProveedor
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Ordenes de Compra por Materia prima consolidado"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "WOrdArtCon.rpt"
            Else
        Listado.ReportFileName = "WOrdPrvCon.rpt"
    End If

    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgOrdartCon.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_WOrden
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArt.SetFocus
    End If
End Sub

Private Sub DesdeArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArt.Text = UCase(DesdeArt.Text)
        Rem HastaArt.Text = DesdeArt.Text
        HastaArt.SetFocus
    End If
End Sub

Private Sub hastaart_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArt.Text = UCase(HastaArt.Text)
        DesdeProv.SetFocus
    End If
End Sub

Private Sub DesdeProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProv.SetFocus
    End If
End Sub

Private Sub HastaProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArt.SetFocus
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Por M.P."
    Tipo.AddItem "Por Proveedor"
    
    Tipo.ListIndex = 0

    Planta.Clear
    
    Planta.AddItem "Consolidado"
    Planta.AddItem "Planta"
    
    Planta.ListIndex = 0
    
    TipoOrden.Clear
    
    TipoOrden.AddItem "Local"
    TipoOrden.AddItem "Importacion"
    TipoOrden.AddItem "Prestamo"
    TipoOrden.AddItem "Envases"
    TipoOrden.AddItem "Total"

    TipoOrden.ListIndex = 4

    Rem With rstEmpresa
    Rem     .Index = "Empresa"
    Rem     .Seek "=", Val(WEmpresa)
    Rem     If .NoMatch = False Then
    Rem         PrgOrdart.Caption = "Listado de Orden de Compra por Materia Prima :  " + !Nombre
    Rem     End If
    Rem End With
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeArt.Text = "  -   -   "
    HastaArt.Text = "  -   -   "
    DesdeProv.Text = ""
    HastaProv.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub


