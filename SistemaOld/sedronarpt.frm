VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSedronarPt 
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
      ItemData        =   "sedronarpt.frx":0000
      Left            =   6840
      List            =   "sedronarpt.frx":0007
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
      MaxLength       =   12
      Mask            =   "AA-#####-###"
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
Attribute VB_Name = "PrgSedronarPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Vector(1000, 10) As String
Private WVectorII(1000, 10) As String
Private Clieventas(1000, 10) As String
Private OrdenCompras(1000, 10) As String

Dim rstTerminado As Recordset
Dim spTerminado As String

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
Dim WTerminado As String
Dim WEntradas As Double
Dim WSalidas As Double
Dim Stock1 As Double
Dim Stock2 As Double
Dim WCompras As Double
Dim WDesde As String
Dim WHasta As String
Dim WFechaord As String
Dim Lugar As Integer
Dim LugarClie As Integer
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
        WTerminado = IngresoDatos.Text
        XCodigo = IngresoDatos.Text
                
        If WTerminado <> "" Then
        
            Rem If WTerminado = "PT-08150-100" Then Stop
            Rem If WTerminado = "PT-25530-100" Then Stop
                
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
            Erase Clieventas
            LugarClie = 0
                    
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
                
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    wstock = rstTerminado!Entradas - rstTerminado!Salidas
                    XXDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
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
                    
            Call Calcula_Ventas
            
            If LugarClie = 0 Then
                LugarClie = 1
            End If
            
            For Ciclo = 1 To LugarClie
                
                XArticulo = "AA-000-000"
                XTerminado = WTerminado
                XRenglon = Str$(Ciclo)
                XInicial = Str$(SumaStock1)
                XComprada = Clieventas(Ciclo, 2)
                XFinal = Str$(SumaStock2)
                XAno = Right$(Desde.Text, 4)
                XPeriodo = Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                XCliente = Clieventas(Ciclo, 1)
                XNroSedronar = ""
                
                If Val(XInicial) <> 0 Or Val(ZFinal) <> 0 Or Val(XVendida) <> 0 Then
                        
                    spCliente = "ConsultaCliente " + "'" + XCliente + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        XNroSedronar = IIf(IsNull(rstCliente!NroSedronar), "0", rstCliente!NroSedronar)
                        rstCliente.Close
                    End If
                        
                    Rem XParam = "'" + XTerminado + "','" _
                    rem              + XRenglon + "','" _
                    rem              + XInicial + "','" _
                    rem              + XComprada + "','" _
                    rem              + XFinal + "','" _
                    rem              + XPeriodo + "','" _
                    rem              + XAno + "','" _
                    rem              + XNroSedronar + "'"
                    Rem
                    Rem spSedronar = "AltaSedronar " + XParam
                    Rem Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZGraba = "N"
                    
                    For CicloII = 1 To LugarVectorII
                    
                        XXCliente = WVectorII(CicloII, 1)
                        XXFechaFactura = WVectorII(CicloII, 2)
                        XXNumeroFactura = WVectorII(CicloII, 3)
                        XXCantidadFactura = WVectorII(CicloII, 4)
                        XXOrdFechaFactura = Right$(XXFechaFactura, 4) + Mid$(XXFechaFactura, 4, 2) + Left$(XXFechaFactura, 2)
                        XXNroRemito = WVectorII(CicloII, 5)
                        XXTipoFactura = "FC"
                        If Val(XXCantidadFactura) <> 0 Then
                            XXTipoOPeracion = "Vta."
                            XXTipoOPeracionII = "Kgs."
                                Else
                            XXTipoOPeracion = ""
                            XXTipoOPeracionII = ""
                        End If
                        XXTerminado = XTerminado

                        If XCliente = XXCliente Then
                        
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Sedronar ("
                            ZSql = ZSql + "Articulo ,"
                            ZSql = ZSql + "Terminado ,"
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
                            ZSql = ZSql + "Descripcion )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + XArticulo + "',"
                            ZSql = ZSql + "'" + XTerminado + "',"
                            ZSql = ZSql + "'" + XRenglon + "',"
                            ZSql = ZSql + "'" + XInicial + "',"
                            ZSql = ZSql + "'" + XComprada + "',"
                            ZSql = ZSql + "'" + XFinal + "',"
                            ZSql = ZSql + "'" + XPeriodo + "',"
                            ZSql = ZSql + "'" + XAno + "',"
                            ZSql = ZSql + "'" + XNroSedronar + "',"
                            ZSql = ZSql + "'" + XXTipoFactura + "',"
                            ZSql = ZSql + "'" + XXFechaFactura + "',"
                            ZSql = ZSql + "'" + XXNumeroFactura + "',"
                            ZSql = ZSql + "'" + XXCantidadFactura + "',"
                            ZSql = ZSql + "'" + XXOrdFechaFactura + "',"
                            ZSql = ZSql + "'" + XXTipoOPeracion + "',"
                            ZSql = ZSql + "'" + XXTipoOPeracionII + "',"
                            ZSql = ZSql + "'" + XXDescripcion + "')"
                                
                            spSedronar = ZSql
                            Set rstSedronar = db.OpenRecordset(spSedronar, dbOpenSnapshot, dbSQLPassThrough)
                            
                            ZGraba = "S"
                        
                        End If
                        
                    Next CicloII
                    
                    If ZGraba = "N" Then
                    
                        XXCliente = WVectorII(CicloII, 1)
                        XXFechaFactura = WVectorII(CicloII, 2)
                        XXNumeroFactura = WVectorII(CicloII, 3)
                        XXCantidadFactura = WVectorII(CicloII, 4)
                        XXOrdFechaFactura = Right$(XXFechaFactura, 4) + Mid$(XXFechaFactura, 4, 2) + Left$(XXFechaFactura, 2)
                        XXNroRemito = WVectorII(CicloII, 5)
                        XXTipoFactura = "FC"
                        XXTipoOPeracion = ""
                        XXTipoOPeracionII = ""
                        XXTerminado = XTerminado
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Sedronar ("
                        ZSql = ZSql + "Articulo ,"
                        ZSql = ZSql + "Terminado ,"
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
                        ZSql = ZSql + "Descripcion )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + XArticulo + "',"
                        ZSql = ZSql + "'" + XTerminado + "',"
                        ZSql = ZSql + "'" + XRenglon + "',"
                        ZSql = ZSql + "'" + XInicial + "',"
                        ZSql = ZSql + "'" + XComprada + "',"
                        ZSql = ZSql + "'" + XFinal + "',"
                        ZSql = ZSql + "'" + XPeriodo + "',"
                        ZSql = ZSql + "'" + XAno + "',"
                        ZSql = ZSql + "'" + XNroSedronar + "',"
                        ZSql = ZSql + "'" + XFechaFactura + "',"
                        ZSql = ZSql + "'" + XNumeroFactura + "',"
                        ZSql = ZSql + "'" + XCantidadFactura + "',"
                        ZSql = ZSql + "'" + XXOrdFechaFactura + "',"
                        ZSql = ZSql + "'" + XXTipoOPeracion + "',"
                        ZSql = ZSql + "'" + XXTipoOPeracionII + "',"
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
            Listado.ReportFileName = "SedrII.rpt"
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
    PrgSedronarPt.Hide
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
            WTerminado = WProducto.Text
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                IngresoDatos.Col = 2
                IngresoDatos.Text = rstTerminado!Descripcion
                rstTerminado.Close
                    Else
                IngresoDatos.Col = 2
                IngresoDatos.Text = ""
            End If
            WProducto.Text = "  -     -   "
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
    IngresoDatos.Text = "Producto"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Descripcion"
    
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Where Terminado.Sedronar = 1"
    ZSql = ZSql + " Order by Terminado.Codigo"
    
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                
                    IngresoDatos.Row = Lugar
                
                    IngresoDatos.Col = 1
                    IngresoDatos.Text = rstTerminado!codigo
                
                    IngresoDatos.Col = 2
                    IngresoDatos.Text = rstTerminado!Descripcion
                
                    .MoveNext
                    
                        Else
                    
                    Exit Do
                
                End If
            
            Loop
        End With
        rstTerminado.Close
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
    Rem             With rstCliente
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
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !codigo
                    IngresaItem = Auxi + "      " + !Descripcion
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !codigo
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
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
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Ingre = "S"
        Lugar = 0
        For A = 1 To 1000
            If Vector(A, 1) = "" And Lugar = 0 Then
                Lugar = A
            End If
            If Vector(A, 1) = rstTerminado!codigo Then
                Ingre = "N"
                Exit For
            End If
        Next A
        If Ingre = "S" Then
            IngresoDatos.Row = Lugar
            Vector(Lugar, 1) = rstTerminado!codigo
            IngresoDatos.Col = 1
            IngresoDatos.Text = rstTerminado!codigo
            WTerminado = rstTerminado!codigo
            IngresoDatos.Col = 2
            IngresoDatos.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Descripcion) - WEspacios
                
                For aa = 1 To da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                        Auxi = !codigo
                        IngresaItem = Auxi + "    " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !codigo
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
    
    rstTerminado.Close
    
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
    
    Rem If WTerminado = "PC-013-100" Then Stop
    
    WEntradas = 0
    WSalidas = 0
    
                
    Rem dada
    Rem PROCESA LAS ESTADISTICAS
    Rem dada
    
    Sql1 = "Select Estadistica.Marca, Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
    Sql4 = " and Estadistica.OrdFecha > " + "'" + WFechaord + "'"
    Sql5 = " and Estadistica.Marca <> " + "'" + "X" + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
    
                    If .EOF = True Then
                        Exit Do
                    End If
        
                    WSalidas = WSalidas + rstEstadistica!Cantidad
        
                    .MoveNext
        
                    If .EOF = True Then
                        Exit Do
                    End If
        
                Loop
            End If
    
        End With

        rstEstadistica.Close

    End If
    
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Hoja.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Hoja.Tipo = " + "'" + "T" + "'"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                ADADA = rstHoja!hoja
                aADADA = rstHoja!Fecha
                asdfaADADA = rstHoja!Fechaord
                
                
                WSalidas = WSalidas + rstHoja!Cantidad
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstHoja.Close
    End If
    
    
    
    
    Rem dada
    Rem PROCESA LAS HOJAS
    Rem dada
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Producto = " + "'" + WTerminado + "'"
    Sql4 = " and Hoja.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Hoja.Renglon = " + "'" + "1" + "'"
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstHoja!realant > 0 Then
                        WEntradas = WEntradas + rstHoja!realant
                            Else
                        WEntradas = WEntradas + rstHoja!Real
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            
            End If
            
        End With
        rstHoja.Close
    End If
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    Rem dada
    
    Sql1 = "Select Movvar.Marca, Movvar.Tipo, Movvar.Terminado, Movvar.Cantidad, Movvar.Fecha, Movvar.Codigo, Movvar.Movi, Movvar.Lote, Movvar.TipoMov, Movvar.Observaciones"
    Sql2 = " FROM Movvar"
    Sql3 = " Where Movvar.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Movvar.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Movvar.Tipo = " + "'" + "T" + "'"
    spMovvar = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTerminado = rstMovvar!terminado
                WCantidad = rstMovvar!Cantidad
                WFecha = rstMovvar!Fecha
                WCodigo = rstMovvar!codigo
                WMovi = rstMovvar!Movi
                WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                If WMovi = "E" Then
                    WEntradas = WEntradas + WCantidad
                        Else
                    WSalidas = WSalidas + WCantidad
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
                
        End With
        rstMovvar.Close
    End If
    
    
    
    
    Rem dada
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNO
    Rem dada
    
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM Guia"
    Sql3 = " Where Guia.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and Guia.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and Guia.Tipo = " + "'" + "T" + "'"
    spMovguia = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
    
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
                Do
                
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    WTerminado = rstMovguia!terminado
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!codigo
                    WMovi = rstMovguia!Movi
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov

                    If WMovi = "E" Then
                        WEntradas = WEntradas + WCantidad
                            Else
                        WSalidas = WSalidas + WCantidad
                    End If
                    
                    .MoveNext
                    
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                Loop
            
            End If
                
        End With
        rstMovguia.Close
    End If
    
    
    
    Rem dada
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    Rem dada
    
    
    Sql1 = "Select *"
    Sql2 = " FROM MovLab"
    Sql3 = " Where MovLab.Terminado = " + "'" + WTerminado + "'"
    Sql4 = " and MovLab.FechaOrd > " + "'" + WFechaord + "'"
    Sql5 = " and MovLab.Tipo = " + "'" + "T" + "'"
    spMovlab = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WTerminado = rstMovlab!terminado
                WCantidad = rstMovlab!Cantidad
                WFecha = rstMovlab!Fecha
                WCodigo = rstMovlab!codigo
                WMovi = rstMovlab!Movi
                
                If WMovi = "E" Then
                    WEntradas = WEntradas + WCantidad
                        Else
                    WSalidas = WSalidas + WCantidad
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        rstMovlab.Close
    End If
    
        
    
    
    
    
        
        
        
        
    
    
End Sub


Private Sub Calcula_Ventas()
                    
    
    Sql1 = "Select Estadistica.Tipo, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Fecha, Estadistica.Numero, Estadistica.Cliente, Estadistica.Lote1, Estadistica.Lote2, Estadistica.Lote3, Estadistica.Lote4, Estadistica.Lote5, Estadistica.Canti1, Estadistica.Canti2, Estadistica.Canti3, Estadistica.Canti4, Estadistica.Canti5, Estadistica.Remito, Estadistica.LoteAdicional"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
    Sql4 = " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    Sql5 = " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then

        With rstEstadistica
            .MoveFirst
            If .NoMatch = False Then
                Do
    
                    If .EOF = True Then
                        Exit Do
                    End If
        
                    WFechaFactura = rstEstadistica!Fecha
                    WNumeroFactura = rstEstadistica!Numero
                    WNroRemito = rstEstadistica!remito
                    WCliente = rstEstadistica!cliente
                    WCantidad = rstEstadistica!Cantidad
                                                            
                    Entra = "S"
                                    
                    For Ciclo = 1 To LugarClie
                        If Clieventas(Ciclo, 1) = WCliente Then
                            Clieventas(Ciclo, 2) = Str$(Val(Clieventas(Ciclo, 2)) + Val(WCantidad))
                            Entra = "N"
                            Exit For
                        End If
                    Next Ciclo
                                    
                    If Entra = "S" Then
                        LugarClie = LugarClie + 1
                        Clieventas(LugarClie, 1) = WCliente
                        Clieventas(LugarClie, 2) = WCantidad
                    End If
                    
                    LugarVectorII = LugarVectorII + 1
                    WVectorII(LugarVectorII, 1) = WCliente
                    WVectorII(LugarVectorII, 2) = WFechaFactura
                    WVectorII(LugarVectorII, 3) = WNumeroFactura
                    WVectorII(LugarVectorII, 4) = WCantidad
                    WVectorII(LugarVectorII, 5) = WNumeroFactura
                    WVectorII(LugarVectorII, 6) = WFechaFactura
                    WVectorII(LugarVectorII, 7) = ""
        
                    .MoveNext
        
                    If .EOF = True Then
                        Exit Do
                    End If
        
                Loop
            End If
    
        End With

        rstEstadistica.Close

    End If
    
End Sub





