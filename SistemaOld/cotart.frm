VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCotart 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cotizaciones por Articulo"
   ClientHeight    =   5910
   ClientLeft      =   1965
   ClientTop       =   1860
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2415
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox Moneda 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Materia Prima"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Materia Prima"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCotart.rpt"
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "cotart.frx":0000
      Left            =   240
      List            =   "cotart.frx":0007
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5760
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6840
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCotart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Vector(3, 5) As String
Private WAuxi As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim Paridad As Double
Dim ParidadII As Double

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    Da = 0
    With rstLiscot
        .Index = "Clave"
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
    
    WAno = Right$(Date$, 4)
    WDia = Mid$(Date$, 4, 2)
    WMes = Left$(Date$, 2)
    XClave = WAno + WMes + WDia

    spCambios = "ConsultaCambioOrdFecha  " + "'" + XClave + "'"
    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
    If rstCambios.RecordCount > 0 Then
        With rstCambios
            .MoveLast
            AA1 = rstCambios!Fecha
            aa2 = rstCambios!OrdFecha
            Paridad = rstCambios!Cambio
            ParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
        End With
        rstCambios.Close
            Else
        Paridad = 1
        ParidadII = 1
    End If
    
    Pasa = 0
    Canti = 0
    WOrden = 0
    
    XParam = "'" + Desde.Text + "','" _
            + Hasta.Text + "'"
    
    spCotiza = "ListaCotizaArticuloDesdeHasta" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    If rstCotiza.RecordCount > 0 Then
            
        With rstCotiza
    
            .MoveFirst
            
            Do
            
                WCotiza = !Cotiza
                WArticulo = !Articulo
                WProveedor = !Proveedor
                WFecha = !Fecha
                
                Select Case Moneda.ListIndex
                    Case 0
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio
                            Case 1
                                WPrecio = !Precio / Paridad
                            Case Else
                                WCoeParidad = ParidadII / Paridad
                                WPrecio = !Precio * WCoeParidad
                        End Select
                    Case 1
                        Select Case !Moneda
                            Case 0
                                WPrecio = !Precio * Paridad
                            Case 1
                                WPrecio = !Precio
                            Case Else
                                WPrecio = !Precio * ParidadII
                        End Select
                    Case Else
                        Select Case !Moneda
                            Case 0
                                WCoeParidad = Paridad / ParidadII
                                WPrecio = !Precio * WCoeParidad
                            Case 1
                                WPrecio = !Precio / ParidadII
                            Case Else
                                WPrecio = !Precio
                        End Select
                End Select
                        
                WCondicion = !Condicion
                WObservaciones = !Observaciones
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase Vector
                    Canti = 0
                End If
                
                If Corte1 <> !Proveedor Or Corte2 <> !Articulo Then
                
                    With rstLiscot
                    
                        Rem If Val(Vector(3, 2)) <> 0 Then
                        Rem     WAuxi = Int(Val(Vector(3, 2)) * 100)
                        Rem             Else
                        Rem     If Val(Vector(2, 2)) <> 0 Then
                        Rem         WAuxi = Int(Val(Vector(2, 2)) * 100)
                        Rem             Else
                        Rem         WAuxi = Int(Val(Vector(1, 2)) * 100)
                        Rem     End If
                        Rem End If
                        Rem
                        Rem Call Ceros(WAuxi, 9)
                        If Vector(3, 2) <> "" Then
                            WAuxi = Vector(3, 5)
                                    Else
                            If Vector(2, 2) <> "" Then
                                WAuxi = Vector(2, 5)
                                    Else
                                WAuxi = Vector(1, 5)
                            End If
                        End If
                        WAuxi = Str$(Val(WAuxi) - 90000000)
                        Call Ceros(WAuxi, 9)
                        
                    
                        For Da = 1 To 3
                        
                            If Vector(Da, 1) <> "" Then
                                .AddNew
                                !Proveedor = Corte1
                                !Articulo = Corte2
                                !Fecha = Vector(Da, 1)
                                !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                !Precio = Val(Vector(Da, 2))
                                !Condicion = Vector(Da, 3)
                                !Observaciones = Vector(Da, 4)
                                !Clave = !Proveedor + !Articulo
                                !Orden = WAuxi + !Proveedor
                             
                                .Update
                            End If
                            
                        Next Da
                            
                    End With
                    
                    Corte1 = !Proveedor
                    Corte2 = !Articulo
                    Erase Vector
                    Canti = 0
                    
                End If
                
                Canti = Canti + 1
                
                If Canti > 3 Then
                    For Da = 1 To 2
                        Vector(Da, 1) = Vector(Da + 1, 1)
                        Vector(Da, 2) = Vector(Da + 1, 2)
                        Vector(Da, 3) = Vector(Da + 1, 3)
                        Vector(Da, 4) = Vector(Da + 1, 4)
                        Vector(Da, 5) = Vector(Da + 1, 5)
                    Next Da
                    Canti = 3
                End If
                
                Vector(Canti, 1) = !Fecha
                Vector(Canti, 2) = Str$(WPrecio)
                Vector(Canti, 3) = !Condicion
                Vector(Canti, 4) = !Observaciones
                Vector(Canti, 5) = !FechaOrd
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
        rstCotiza.Close
    
    End If
    
    If Pasa <> 0 Then
        With rstLiscot
                
            Rem If Val(Vector(3, 2)) <> 0 Then
            Rem     WAuxi = Int(Val(Vector(3, 2)) * 100)
            Rem             Else
            Rem     If Val(Vector(2, 2)) <> 0 Then
            Rem         WAuxi = Int(Val(Vector(2, 2)) * 100)
            Rem             Else
            Rem         WAuxi = Int(Val(Vector(1, 2)) * 100)
            Rem     End If
            Rem End If
            Rem
            Rem Call Ceros(WAuxi, 9)
            
            If Vector(3, 5) <> "" Then
                WAuxi = Vector(3, 5)
                        Else
                If Vector(2, 5) <> "" Then
                    WAuxi = Vector(2, 5)
                        Else
                    WAuxi = Vector(1, 5)
                End If
            End If
            WAuxi = Str$(Val(WAuxi) - 90000000)
            Call Ceros(WAuxi, 9)
                
            For Da = 1 To 3
                    
                If Vector(Da, 1) <> "" Then
                    .AddNew
                    !Proveedor = Corte1
                    !Articulo = Corte2
                    !Fecha = Vector(Da, 1)
                    !FechaOrd = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                    !Precio = Val(Vector(Da, 2))
                    !Condicion = Vector(Da, 3)
                    !Observaciones = Vector(Da, 4)
                    !Clave = !Proveedor + !Articulo
                    !Orden = WAuxi + !Proveedor
                    .Update
                End If
                
            Next Da
                        
        End With
    End If
    
    Da = 0
    With rstLiscot
        .Index = "Clave"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WDescriProveedor = ""
                WArticulo = !Articulo
                WDescriArticulo = ""
                    
                WCategoriaI = ""
                WCategoriaII = ""
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                spProveedor = ZSql
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                
                    WDescriProveedor = RstProveedor!Nombre
                    
                    ZCategoriaI = IIf(IsNull(RstProveedor!CategoriaI), "0", RstProveedor!CategoriaI)
                    ZCategoriaII = IIf(IsNull(RstProveedor!CategoriaII), "0", RstProveedor!CategoriaII)
        
                    If ZCategoriaI = 1 Then
                        WCategoriaI = "A"
                            Else
                        If ZCategoriaI = 2 Then
                            WCategoriaI = "B"
                                Else
                            If ZCategoriaI = 3 Then
                                WCategoriaI = "C"
                                    Else
                                If ZCategoriaI = 4 Then
                                    WCategoriaI = "E"
                                End If
                            End If
                        End If
                    End If
                    
                    WCategoriaII = "S/C"
                    If ZCategoriaII = 1 Then
                        WCategoriaII = "Muy Bueno"
                            Else
                        If ZCategoriaII = 2 Then
                            WCategoriaII = "Bueno"
                                Else
                            If ZCategoriaII = 3 Then
                                WCategoriaII = "Regular"
                                    Else
                                If ZCategoriaII = 4 Then
                                    WCategoriaII = "Malo"
                                End If
                            End If
                        End If
                    End If
                    
                    RstProveedor.Close
                End If
                
                ZZIngre = ""
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                    rstArticulo.Close
                End If
                
                If ZZTipoMp = 1 Then
                
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            WEmpresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0004"
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Homologa"
                    ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
                    ZSql = ZSql + " and CodigoMp = " + "'" + WArticulo + "'"
                    ZSql = ZSql + " and Estado = " + "'" + "1" + "'"
                    spHomologa = ZSql
                    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHomologa.RecordCount > 0 Then
                        ZZIngre = "  (H)   "
                        rstHomologa.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    If ZZIngre = "" Then
                    
                        XEmpresa = WEmpresa
                        Select Case Val(WEmpresa)
                            Case 1, 3, 5, 6, 7, 10, 11
                                WEmpresa = "0001"
                                txtOdbc = "Empresa01"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0003"
                                txtOdbc = "Empresa03"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0005"
                                txtOdbc = "Empresa05"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0006"
                                txtOdbc = "Empresa06"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0007"
                                txtOdbc = "Empresa07"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0010"
                                txtOdbc = "Empresa10"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0011"
                                txtOdbc = "Empresa11"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                            Case Else
                                WEmpresa = "0002"
                                txtOdbc = "Empresa02"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0004"
                                txtOdbc = "Empresa04"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0008"
                                txtOdbc = "Empresa08"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                                
                                WEmpresa = "0009"
                                txtOdbc = "Empresa09"
                                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                                
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Orden"
                                ZSql = ZSql + " Where Articulo = " + "'" + WArticulo + "'"
                                ZSql = ZSql + " and Proveedor = " + "'" + WProveedor + "'"
                                spOrden = ZSql
                                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                                If rstOrden.RecordCount > 0 Then
                                    ZZIngre = "  (H)   "
                                    rstOrden.Close
                                End If
                            
                        End Select
                    
                        Call Conecta_Empresa
                        
                    End If
                    
                End If
                
                WDescriProveedor = ZZIngre + Trim(WDescriProveedor)
                    
                If WCategoriaI <> "" And WCategoriaII <> "" Then
                    WDescriProveedor = Trim(WDescriProveedor) + " (" + WCategoriaI + " - " + WCategoriaII + ")"
                End If
                
                !DescriProveedor = Left$(WDescriProveedor, 50)
                !DescriArticulo = WDescriArticulo
                
                Select Case Moneda.ListIndex
                    Case 0
                        !Titulo = "(En Dolares)"
                    Case 1
                        !Titulo = "(En Pesos)"
                    Case Else
                        !Titulo = "(En Euros)"
                End Select
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Cotizaciones por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Listcot.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstLiscot
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
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
    
    
    Desde.SetFocus
    PrgCotart.Hide
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
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Liscot
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    XEmpresa = WEmpresa
        
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)


    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCotart.Caption = "Listado de Cotizaciones por Producto"
        End If
    End With
    
    Moneda.Clear
    
    Moneda.AddItem "Dolares"
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Euros"
    
    Moneda.ListIndex = 0
    
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Impresora.Value = False
    Panta.Value = True
    Frame2.Visible = True
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
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
    
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                Desde.Text = rstArticulo!Codigo
                Hasta.Text = rstArticulo!Codigo
            End If
            Desde.SetFocus
    End Select
    
End Sub

