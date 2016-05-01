VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSedronar 
   AutoRedraw      =   -1  'True
   Caption         =   "Declaracion Jurada (Sedronar)"
   ClientHeight    =   7365
   ClientLeft      =   450
   ClientTop       =   825
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11100
   Begin MSFlexGridLib.MSFlexGrid IngresoDatos 
      Height          =   2535
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      _Version        =   327680
      Rows            =   1000
      Cols            =   3
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      Height          =   300
      Left            =   480
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   4335
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
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
         TabIndex        =   1
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FechaAnterior 
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
      Begin VB.Label Label4 
         Caption         =   "Cierre Anterior"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Sedronar.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
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
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   6495
      ItemData        =   "sedronar.frx":0000
      Left            =   6840
      List            =   "sedronar.frx":0007
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox WProducto 
      Height          =   300
      Left            =   2640
      TabIndex        =   18
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
      Caption         =   "Ingreso de Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   5640
      Width           =   1935
   End
End
Attribute VB_Name = "PrgSedronar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Vector(1000, 10) As String
Private WVectorII(1000, 10) As String
Private ProveCompras(1000, 10) As String
Private OrdenCompras(1000, 10) As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstSedronar As Recordset
Dim spSedronar As String
Dim XParam As String
Dim WArticulo As String
Dim WEntradas As Double
Dim WSalidas As Double
Dim Stock1 As Double
Dim Stock2 As Double
Dim WCompras As Double
Dim WDesde As String
Dim WHasta As String
Dim WFechaord As String
Dim Lugar As Integer
Dim LugarProve As Integer
Dim LugarOrden As Integer
Dim WEmpre(10) As String
Dim LugarVectorII As Integer

Private Sub Acepta_Click()

    For A = 1 To 999
        With rstSedro
            .Index = "Clave"
            .Seek "=", A
            If .NoMatch = False Then
                .Delete
            End If
        End With
    Next A

    Lugar = 0
    For A = 1 To 999
        WProd = Vector(A, 1)
        If WProd <> "" Then
            Lugar = Lugar + 1
            With rstSedro
                .AddNew
                !Clave = Lugar
                !Producto = WProd
                .Update
            End With
        End If
    Next A

    spSedronar = "BorrarSedronar "
    Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Declaracion Jurada (Sedronar)"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    EmpresaAnterior = WEmpresa
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7
            DesdeEmpresa = 1
            HastaEmpresa = 7
            WEmpre(1) = "0001"
            WEmpre(2) = "0003"
            WEmpre(3) = "0005"
            WEmpre(4) = "0006"
            WEmpre(5) = "0007"
            WEmpre(6) = "0010"
            WEmpre(7) = "0011"
        Case Else
            DesdeEmpresa = 1
            HastaEmpresa = 4
            WEmpre(1) = "0002"
            WEmpre(2) = "0004"
            WEmpre(3) = "0008"
            WEmpre(4) = "0009"
    End Select
    
    For A = 1 To 999
    
        iRow = A
    
        IngresoDatos.Col = 1
        IngresoDatos.Row = iRow
        WArticulo = IngresoDatos.Text
        XCodigo = IngresoDatos.Text
        XXDescripcion = ""
                
        If WArticulo <> "" Then
                
            WAno = Right$(Desde.Text, 4)
            WMes = Mid$(Desde.Text, 4, 2)
            WDia = Left$(Desde.Text, 2)
            WFechaord = WAno + WMes + WDia
                
            WEntradas = 0
            WSalidas = 0
            SumaStock1 = 0
            SumaStock2 = 0
                    
            Erase WVectorII
            LugarVectorII = 0
            Erase ProveCompras
            LugarProve = 0
                    
            For CiclaEmpre = DesdeEmpresa To HastaEmpresa
            
                Select Case Val(WEmpre(CiclaEmpre))
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
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    wstock = rstArticulo!Entradas - rstArticulo!Salidas
                    XXDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            
                WEntradas = 0
                WSalidas = 0
                        
                WAno = Right$(FechaAnterior.Text, 4)
                WMes = Mid$(FechaAnterior.Text, 4, 2)
                WDia = Left$(FechaAnterior.Text, 2)
                WFechaord = WAno + WMes + WDia
        
                Call calcula_datos
                Stock1 = wstock - WEntradas + WSalidas
                Call Redondeo(Stock1)
                SumaStock1 = SumaStock1 + Stock1
                
                WEntradas = 0
                WSalidas = 0
                
                WAno = Right$(Hasta.Text, 4)
                WMes = Mid$(Hasta.Text, 4, 2)
                WDia = Left$(Hasta.Text, 2)
                WFechaord = WAno + WMes + WDia
        
                Call calcula_datos
                Stock2 = wstock - WEntradas + WSalidas
                Call Redondeo(Stock2)
                SumaStock2 = SumaStock2 + Stock2
                        
                WAno = Right$(Desde.Text, 4)
                WMes = Mid$(Desde.Text, 4, 2)
                WDia = Left$(Desde.Text, 2)
                WDesde = WAno + WMes + WDia
                        
                WAno = Right$(Hasta.Text, 4)
                WMes = Mid$(Hasta.Text, 4, 2)
                WDia = Left$(Hasta.Text, 2)
                WHasta = WAno + WMes + WDia
                
                Call calcula_Compras
                        
            Next CiclaEmpre
                    
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
                    
            If LugarProve = 0 Then
                LugarProve = 1
            End If
            
            Rem DADA
            Rem DADA
            Rem DADA
            Rem esta es la posta
            Rem If WArticulo = "PC-013-100" Then Stop
                    
            For Ciclo = 1 To LugarProve
                
                XArticulo = WArticulo
                XRenglon = Str$(Ciclo)
                XInicial = Str$(SumaStock1)
                XComprada = ProveCompras(Ciclo, 2)
                XFinal = Str$(SumaStock2)
                XAno = Right$(Desde.Text, 4)
                XPeriodo = Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                XProve = ProveCompras(Ciclo, 1)
                XNroInsc = ""
                
                If Val(XInicial) <> 0 Or Val(ZFinal) <> 0 Or Val(XComprada) <> 0 Then
                        
                    spProveedor = "ConsultaProveedores " + "'" + XProve + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        XNroInsc = IIf(IsNull(RstProveedor!NroInsc), "0", RstProveedor!NroInsc)
                        RstProveedor.Close
                    End If
                        
                    Rem XParam = "'" + XArticulo + "','" _
                    rem              + XRenglon + "','" _
                    rem              + XInicial + "','" _
                    rem              + XComprada + "','" _
                    rem              + XFinal + "','" _
                    rem              + XPeriodo + "','" _
                    rem              + XAno + "','" _
                    rem              + XNroInsc + "'"
                    Rem
                    Rem spSedronar = "AltaSedronar " + XParam
                    Rem Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZGraba = "N"
                    
                    For CicloII = 1 To LugarVectorII
                    
                        XXProve = WVectorII(CicloII, 1)
                        XXFechaFactura = WVectorII(CicloII, 2)
                        XXNumeroFactura = WVectorII(CicloII, 3)
                        XXCantidadFactura = WVectorII(CicloII, 4)
                        XXOrdFechaFactura = Right$(XXFechaFactura, 4) + Mid$(XXFechaFactura, 4, 2) + Left$(XXFechaFactura, 2)
                        XXNroRemito = WVectorII(CicloII, 5)
                        XXFechaOrden = WVectorII(CicloII, 6)
                        XXOrdFechaOrden = Right$(XXFechaOrden, 4) + Mid$(XXFechaOrden, 4, 2) + Left$(XXFechaOrden, 2)
                        XXOrden = WVectorII(CicloII, 7)
                        XXTipoFactura = "FC"
                        If Val(XXCantidadFactura) <> 0 Then
                            XXTipoOPeracion = "Cpa"
                            XXTipoOPeracionII = "Kgs."
                                Else
                            XXTipoOPeracion = ""
                            XXTipoOPeracionII = ""
                        End If
                        XXTerminado = ""
                        
                        If Trim(XXNumeroFactura) = "" Then
                            XXFechaFactura = XXFechaOrden
                            XXOrdFechaFactura = XXOrdFechaOrden
                            XXNumeroFactura = XXOrden
                            XXTipoFactura = "OC"
                        End If

                        If XProve = XXProve Then
                        
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Sedronar ("
                            ZSql = ZSql + "Articulo ,"
                            ZSql = ZSql + "Renglon ,"
                            ZSql = ZSql + "Inicial ,"
                            ZSql = ZSql + "Comprada ,"
                            ZSql = ZSql + "Final ,"
                            ZSql = ZSql + "Periodo ,"
                            ZSql = ZSql + "Ano ,"
                            ZSql = ZSql + "NroInsc ,"
                            ZSql = ZSql + "TipoFactura ,"
                            ZSql = ZSql + "FechaFactura ,"
                            ZSql = ZSql + "NumeroFactura ,"
                            ZSql = ZSql + "CantidadFactura ,"
                            ZSql = ZSql + "OrdFechaFactura ,"
                            ZSql = ZSql + "TipoOperacion ,"
                            ZSql = ZSql + "TipoOperacionII ,"
                            ZSql = ZSql + "Terminado ,"
                            ZSql = ZSql + "Descripcion )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + XArticulo + "',"
                            ZSql = ZSql + "'" + XRenglon + "',"
                            ZSql = ZSql + "'" + XInicial + "',"
                            ZSql = ZSql + "'" + XComprada + "',"
                            ZSql = ZSql + "'" + XFinal + "',"
                            ZSql = ZSql + "'" + XPeriodo + "',"
                            ZSql = ZSql + "'" + XAno + "',"
                            ZSql = ZSql + "'" + XNroInsc + "',"
                            ZSql = ZSql + "'" + XXTipoFactura + "',"
                            ZSql = ZSql + "'" + XXFechaFactura + "',"
                            ZSql = ZSql + "'" + XXNumeroFactura + "',"
                            ZSql = ZSql + "'" + XXCantidadFactura + "',"
                            ZSql = ZSql + "'" + XXOrdFechaFactura + "',"
                            ZSql = ZSql + "'" + XXTipoOPeracion + "',"
                            ZSql = ZSql + "'" + XXTipoOPeracionII + "',"
                            ZSql = ZSql + "'" + XXTerminado + "',"
                            ZSql = ZSql + "'" + XXDescripcion + "')"
                                
                            spSedronar = ZSql
                            Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZGraba = "S"
                        
                        End If
                        
                    Next CicloII
                    
                    If ZGraba = "N" Then
                    
                        XXProve = WVectorII(CicloII, 1)
                        XXFechaFactura = WVectorII(CicloII, 2)
                        XXNumeroFactura = WVectorII(CicloII, 3)
                        XXCantidadFactura = WVectorII(CicloII, 4)
                        XXOrdFechaFactura = Right$(XXFechaFactura, 4) + Mid$(XXFechaFactura, 4, 2) + Left$(XXFechaFactura, 2)
                        XXNroRemito = WVectorII(CicloII, 5)
                        XXFechaOrden = WVectorII(CicloII, 6)
                        XXOrdFechaOrden = Right$(XXFechaOrden, 4) + Mid$(XXFechaOrden, 4, 2) + Left$(XXFechaOrden, 2)
                        XXOrden = WVectorII(CicloII, 7)
                        XXTipoFactura = ""
                        XXTipoOPeracion = ""
                        XXTipoOPeracionII = ""
                        XXTerminado = ""
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Sedronar ("
                        ZSql = ZSql + "Articulo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Inicial ,"
                        ZSql = ZSql + "Comprada ,"
                        ZSql = ZSql + "Final ,"
                        ZSql = ZSql + "Periodo ,"
                        ZSql = ZSql + "Ano ,"
                        ZSql = ZSql + "NroInsc ,"
                        ZSql = ZSql + "FechaFactura ,"
                        ZSql = ZSql + "NumeroFactura ,"
                        ZSql = ZSql + "CantidadFactura ,"
                        ZSql = ZSql + "OrdFechaFactura ,"
                        ZSql = ZSql + "TipoOperacion ,"
                        ZSql = ZSql + "TipoOperacionII ,"
                        ZSql = ZSql + "Terminado ,"
                        ZSql = ZSql + "Descripcion )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + XArticulo + "',"
                        ZSql = ZSql + "'" + XRenglon + "',"
                        ZSql = ZSql + "'" + XInicial + "',"
                        ZSql = ZSql + "'" + XComprada + "',"
                        ZSql = ZSql + "'" + XFinal + "',"
                        ZSql = ZSql + "'" + XPeriodo + "',"
                        ZSql = ZSql + "'" + XAno + "',"
                        ZSql = ZSql + "'" + XNroInsc + "',"
                        ZSql = ZSql + "'" + XFechaFactura + "',"
                        ZSql = ZSql + "'" + XNumeroFactura + "',"
                        ZSql = ZSql + "'" + XCantidadFactura + "',"
                        ZSql = ZSql + "'" + XXOrdFechaFactura + "',"
                        ZSql = ZSql + "'" + XXTipoOPeracion + "',"
                        ZSql = ZSql + "'" + XXTipoOPeracionII + "',"
                        ZSql = ZSql + "'" + XXTerminado + "',"
                        ZSql = ZSql + "'" + XXDescripcion + "')"
                            
                        spSedronar = ZSql
                        Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)
                    
                    End If
                
                End If
                        
            Next Ciclo
                
        End If
        
DADA:
            
    Next A

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7
            Listado.ReportFileName = "Sedronar.rpt"
        Case Else
            Listado.ReportFileName = "SedroII.rpt"
    End Select
    
    Listado.ReportFileName = "Sedronar.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Sedronar.Articulo, Sedronar.Inicial, Sedronar.Comprada, Sedronar.Final, Sedronar.Periodo, Sedronar.Ano, Sedronar.NroInsc, Sedronar.FechaFactura, Sedronar.NumeroFactura, Sedronar.CantidadFactura, Sedronar.OrdFechaFactura, Sedronar.TipoFactura, Sedronar.TipoOperacion,  Sedronar.Descripcion, Sedronar.TipoOperacionII,   " _
                        + "From " _
                        + DSQ + ".dbo.Sedronar Sedronar " _
                        + "Where " _
                        + "Sedronar.Articulo >= 'AA-000-000' AND Sedronar.Articulo <= 'ZZ-999-999'"
    
    Rem Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    WProducto.SetFocus
    
End Sub

Private Sub Cancela_click()

    For A = 1 To 999
        With rstSedro
            .Index = "Clave"
            .Seek "=", A
            If .NoMatch = False Then
                .Delete
            End If
        End With
    Next A

    Lugar = 0
    For A = 1 To 999
        WProd = Vector(A, 1)
        If WProd <> "" Then
            Lugar = Lugar + 1
            With rstSedro
                .AddNew
                !Clave = Lugar
                !Producto = WProd
                .Update
            End With
        End If
    Next A

    With rstEmpresa
        .Close
    End With
    PrgSedronar.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Sedro
    OPEN_FILE_Empresa
End Sub

Private Sub IngresoDatos_DblClick()
    IngresoDatos.Col = 1
    IngresoDatos.Text = ""
    IngresoDatos.Col = 2
    IngresoDatos.Text = ""
    Lugar = IngresoDatos.Row
    Vector(Lugar, 1) = ""
    WProducto.SetFocus
End Sub

Private Sub WProducto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WProducto.Text = UCase(WProducto.Text)
    
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = WProducto.Text Then
                Ingre = "N"
                Exit For
            End If
        Next A
                            
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = WProducto.Text
            IngresoDatos.Col = 1
            IngresoDatos.Text = WProducto.Text
            WArticulo = WProducto.Text
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                IngresoDatos.Col = 2
                IngresoDatos.Text = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                IngresoDatos.Col = 2
                IngresoDatos.Text = ""
            End If
            WProducto.Text = "  -   -   "
            WProducto.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()

    IngresoDatos.Clear
    Erase Vector
    
    IngresoDatos.ColWidth(0) = 150
    IngresoDatos.ColWidth(1) = 1600
    IngresoDatos.ColWidth(2) = 3500
    
    IngresoDatos.Row = 0
    
    IngresoDatos.Col = 1
    IngresoDatos.Text = "Articulo"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Descripcion"
    
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Sedronar = 1"
    ZSql = ZSql + " Order by Articulo.Codigo"
    
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                
                    IngresoDatos.Row = Lugar
                
                    IngresoDatos.Col = 1
                    IngresoDatos.Text = rstArticulo!Codigo
                
                    IngresoDatos.Col = 2
                    IngresoDatos.Text = rstArticulo!Descripcion
                
                    .MoveNext
                    
                        Else
                    
                    Exit Do
                
                End If
            
            Loop
        End With
        rstArticulo.Close
    End If
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    FechaAnterior.Text = "  /  /    "
    
End Sub

Private Sub Proceso_Click()

    Rem With rstProceso1
    Rem     .Index = "Numero"
    Rem     .Seek ">=", da
    Rem     If .NoMatch = False Then
    Rem         Do
    Rem
    Rem             WProveedor = !Proveedor
    Rem
    Rem             With rstProveedor
    Rem                 .Index = "Proveedor"
    Rem                 .Seek "=", WProveedor
    Rem                 If .NoMatch = False Then
    Rem                     WNombre = !Nombre
    Rem                 End If
    Rem             End With
    Rem
    Rem
    Rem             Lugar1 = Int(!Numero / 10)
    Rem             Lugar2 = !Numero - Lugar1
    Rem
    Rem             DBGrid1.FirstRow = Lugar1
    Rem             DBGrid1.Row = Lugar2 - 1
    Rem             DBGrid1.Col = 0
    Rem             DBGrid1.Text = WProveedor
    Rem             DBGrid1.Col = 1
    Rem             DBGrid1.Text = WNombre
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End If
    Rem End With

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1

End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Codigo
                    IngresaItem = Auxi + "      " + !Descripcion
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Codigo
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = rstArticulo!Codigo Then
                Ingre = "N"
                Exit For
            End If
        Next A
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = rstArticulo!Codigo
            IngresoDatos.Col = 1
            IngresoDatos.Text = rstArticulo!Codigo
            WArticulo = rstArticulo!Codigo
            IngresoDatos.Col = 2
            IngresoDatos.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Descripcion) - WEspacios
                
                For aa = 1 To da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                        Auxi = !Codigo
                        IngresaItem = Auxi + "    " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Codigo
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
    
    rstArticulo.Close
    
    End If
    
    End If

End Sub

Private Sub FechaAnterior_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WProducto.SetFocus
    End If
End Sub

Private Sub calcula_datos()

    Rem PROCESA LOS LAUDOS
    
    Rem If WArticulo = "PC-013-100" Then Stop
    
    WEntradas = 0
    WSalidas = 0
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spLaudo = "ListaLaudoArticuloDesdeHastaFecha " + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstLaudo!Fecha, 4)
                Rem WMes = Mid$(rstLaudo!Fecha, 4, 2)
                Rem WDia = Left$(rstLaudo!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                
                Rem If WCompara > WFechaord Then
                Rem     If rstLaudo!Articulo = WArticulo Then
                        WLiberada = IIf(IsNull(rstLaudo!Liberadaant), 0, rstLaudo!Liberadaant)
                        If WLiberada = 0 Then
                            WLiberada = rstLaudo!Liberada
                        End If
                        WEntradas = WEntradas + WLiberada
                Rem     End If
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spHoja = "ListaHojaArticuloDesdeHastaFecha" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstHoja!Fecha, 4)
                Rem WMes = Mid$(rstHoja!Fecha, 4, 2)
                Rem WDia = Left$(rstHoja!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                Rem         XX = rstHoja!Clave
                        Rem WCantidad = rstHoja!Canti1 + rstHoja!Canti2 + rstHoja!Canti3
                        Rem If WCantidad = 0 Then
                            WCantidad = rstHoja!Cantidad
                        Rem End If
                        WSalidas = WSalidas + WCantidad
                Rem     End If
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstHoja!Articulo > WArticulo Then
                Rem     Exit Do
                Rem End If
                
            Loop
            End If
        
        End With
        
        rstHoja.Close
        
    End If
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spMovvar = "ListaMovvarArticuloDesdeHastaFecha" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovvar!Fecha, 4)
                Rem WMes = Mid$(rstMovvar!Fecha, 4, 2)
                Rem WDia = Left$(rstMovvar!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                    Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                Rem     End If
                Rem End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        
        rstMovvar.Close
        
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHastaFecha" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovguia!Fecha, 4)
                Rem WMes = Mid$(rstMovguia!Fecha, 4, 2)
                Rem WDia = Left$(rstMovguia!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                        WCantidad = IIf(IsNull(rstMovguia!Cantidadant), 0, rstMovguia!Cantidadant)
                        If WCantidad = 0 Then
                            WCantidad = rstMovguia!Cantidad
                        End If
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + WCantidad
                                Else
                            WSalidas = WSalidas + WCantidad
                        End If
                Rem     End If
                Rem End If
                
                .MoveNext
            
                If .EOF = True Then
                    Exit Do
                End If
                                                                            
            Loop
            End If
            
        End With
        
        rstMovguia.Close
        
    End If
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHastaFecha" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovlab!Fecha, 4)
                Rem WMes = Mid$(rstMovlab!Fecha, 4, 2)
                Rem WDia = Left$(rstMovlab!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                        WCantidad = rstMovlab!Cantidad
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + WCantidad
                                Else
                            WSalidas = WSalidas + WCantidad
                        End If
                Rem     End If
                Rem End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    
    Rem PROCESA LAS VENTAS
    
    If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHastaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
                Rem     WAno = Right$(rstEstadistica!Fecha, 4)
                Rem     WMes = Mid$(rstEstadistica!Fecha, 4, 2)
                Rem     WDia = Left$(rstEstadistica!Fecha, 2)
                Rem     WCompara = WAno + WMes + WDia
                        
                Rem     If WCompara > WFechaord Then
                        If rstEstadistica!Tipo = 1 Then
                            WSalidas = WSalidas + rstEstadistica!Cantidad
                                Else
                            WEntradas = WEntradas + rstEstadistica!Cantidad
                        End If
                Rem     End If
                Rem End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    End If
    
End Sub


Private Sub calcula_Compras()
                    
    Erase OrdenCompras
    LugarOrden = 0
    
    Rem If WArticulo = "PC-013-100" Then Stop

    Rem PROCESA LOS LAUDOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoArticuloDesdeHasta " + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)

                WAno = Right$(rstLaudo!Fecha, 4)
                WMes = Mid$(rstLaudo!Fecha, 4, 2)
                WDia = Left$(rstLaudo!Fecha, 2)
                WCompara = WAno + WMes + WDia
                        
                If WCompara >= WDesde And WCompara <= WHasta Then
                    If rstLaudo!Articulo = WArticulo Then
                    
                        If WLiberadaAnt <> 0 Then
                            WSuma = WLiberadaAnt
                                Else
                            WSuma = WLiberada
                        End If
                    
                        LugarOrden = LugarOrden + 1
                        OrdenCompras(LugarOrden, 1) = rstLaudo!Orden
                        OrdenCompras(LugarOrden, 2) = Str$(WSuma)
                        OrdenCompras(LugarOrden, 3) = rstLaudo!Articulo
                        OrdenCompras(LugarOrden, 4) = rstLaudo!informe
                        OrdenCompras(LugarOrden, 5) = WEmpresa
                    
                    End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Tipo = 4"
    ZSql = ZSql + " and Orden.Articulo = " + "'" + WArticulo + "'"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then

        With rstOrden
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    WAno = Right$(rstOrden!Fecha, 4)
                    WMes = Mid$(rstOrden!Fecha, 4, 2)
                    WDia = Left$(rstOrden!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                            
                    If WCompara >= WDesde And WCompara <= WHasta Then
                        If rstOrden!Articulo = WArticulo Then
                            
                            LugarOrden = LugarOrden + 1
                            OrdenCompras(LugarOrden, 1) = rstOrden!Orden
                            OrdenCompras(LugarOrden, 2) = Str$(rstOrden!Cantidad)
                            OrdenCompras(LugarOrden, 3) = rstOrden!Articulo
                            OrdenCompras(LugarOrden, 4) = ""
                            OrdenCompras(LugarOrden, 5) = WEmpresa
                            
                        End If
                    End If
                        
                    .MoveNext
            
                    If .EOF = True Then
                        Exit Do
                    End If
                                                                            
                Loop
            End If
            
        End With
                
        rstOrden.Close
    End If
    
    For CicloProve = 1 To LugarOrden
    
        WOrden = OrdenCompras(CicloProve, 1)
        WCantidad = OrdenCompras(CicloProve, 2)
        WArticulo = OrdenCompras(CicloProve, 3)
        WInforme = OrdenCompras(CicloProve, 4)
        XXEmpresa = OrdenCompras(CicloProve, 5)
        
        WFechaFactura = ""
        WNumeroFactura = ""
        WNroRemito = ""
        WProve = ""
        WFechaOrden = ""
        
        Rem If WArticulo = "PC-013-100" Then Stop
        
        XEmpresa = WEmpresa
        Select Case Val(XXEmpresa)
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
    
        spOrden = "ListaOrden " + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProve = rstOrden!proveedor
            WFechaOrden = rstOrden!Fecha
            rstOrden.Close
        End If
                                                
        If Trim(WInforme) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Informe = " + "'" + WInforme + "'"
            spinforme = ZSql
            Set rstInforme = db.OpenRecordset(spinforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                WNroRemito = rstInforme!remito
                rstInforme.Close
            End If
        End If
        
        Auxi = WNroRemito
        Auxi = Trim(Auxi)
        
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
        
        
        XEmpresa = WEmpresa
        
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
        End Select
                                                
        If Trim(WNroRemito) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ivacomp"
            ZSql = ZSql + " Where Ivacomp.Remito LIKE " + "'" + "%" + Auxi + "%" + "'"
            ZSql = ZSql + " Order by Ivacomp.Proveedor,Ivacomp.OrdFecha"
            Rem ZSql = ZSql + " and Ivacomp.Proveedor = " + "'" + WProve + "'"
            spIvacomp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            If rstIvaComp.RecordCount > 0 Then
            
                With rstIvaComp
            
                    .MoveFirst
                    
                    If .NoMatch = False Then
                        Do
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            If WProve = rstIvaComp!proveedor Then
                                WFechaFactura = rstIvaComp!Fecha
                                WNumeroFactura = rstIvaComp!Numero
                            End If
                                
                            .MoveNext
                    
                            If .EOF = True Then
                                Exit Do
                            End If
                                                                                    
                        Loop
                    End If
                    
                End With
            
                rstIvaComp.Close
            End If
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
                                                
        Entra = "S"
                        
        For Ciclo = 1 To LugarProve
            If ProveCompras(Ciclo, 1) = WProve Then
                ProveCompras(Ciclo, 2) = Str$(Val(ProveCompras(Ciclo, 2)) + Val(WCantidad))
                Entra = "N"
                Exit For
            End If
        Next Ciclo
                        
        If Entra = "S" Then
            LugarProve = LugarProve + 1
            ProveCompras(LugarProve, 1) = WProve
            ProveCompras(LugarProve, 2) = WCantidad
        End If
        
        LugarVectorII = LugarVectorII + 1
        WVectorII(LugarVectorII, 1) = WProve
        WVectorII(LugarVectorII, 2) = WFechaFactura
        WVectorII(LugarVectorII, 3) = WNumeroFactura
        WVectorII(LugarVectorII, 4) = WCantidad
        WVectorII(LugarVectorII, 5) = WNroRemito
        WVectorII(LugarVectorII, 6) = WFechaOrden
        WVectorII(LugarVectorII, 7) = WOrden
        
    Next CicloProve
    
End Sub





