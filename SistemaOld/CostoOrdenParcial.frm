VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCostoOrdenParcial 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Calculo de Costo de Nacionalizacion de Mercaderia"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2055
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   4815
      Begin VB.TextBox Carpeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Codigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Movimiento"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCarpetaParcial.rpt"
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
      ItemData        =   "CostoOrdenParcial.frx":0000
      Left            =   840
      List            =   "CostoOrdenParcial.frx":0007
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
Attribute VB_Name = "PrgCostoOrdenParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim Vector(100, 20) As String
Dim Gastos(100, 10) As String
Dim rstCarpeta As Recordset
Dim spCarpeta As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstMovGasParcial As Recordset
Dim spMovGasParcial As String
Dim rstMovGasParcialArti As Recordset
Dim spMovGasParcialArti As String
Dim WArancel As String
Dim WCarpeta As String
Dim WCosto As Double
Dim WLeyenda As Integer
Dim CargaEmpresa(12, 2) As String

Private Sub Acepta_Click()

    WCarpeta = Carpeta.Text
    Call Ceros(WCarpeta, 6)

    spCarpeta = "BorrarCarpeta"
    Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)

    Renglon = 0
    Erase Gastos
    ImpoGastos = 0
    ImpoSeguro = 0
    ImpoFlete = 0
    
    spMovgas = "ListaMovgas " + "'" + Carpeta.Text + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstMovgas!Concepto <> 10 Then
                        WArancel = Str$(rstMovgas!Derechos)
                        WOrden = Str$(rstMovgas!Orden)
                        WEmpresaOtro = rstMovgas!Empresa
                        Select Case rstMovgas!Concepto
                            Case 2
                                ImpoSeguro = ImpoSeguro + rstMovgas!Importe
                            Case 4, 5
                                ImpoFlete = ImpoFlete + rstMovgas!Importe
                            Case Else
                                Renglon = Renglon + 1
                                Gastos(Renglon, 1) = Str$(rstMovgas!Concepto)
                                Gastos(Renglon, 2) = Str$(rstMovgas!Importe)
                                ImpoGastos = ImpoGastos + rstMovgas!Importe
                        End Select
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovgas.Close
                
    End If
    
    WGastos = Renglon
    ImpoArancel = 0
    WTotal = 0
    
    Renglon = 0
    Erase Vector
    
    EmpresaAnterior = WEmpresa
    Select Case WEmpresaOtro
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
    
    WTotalImpo = 0
    WTotalPeso = 0
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    
                    Vector(Renglon, 1) = Carpeta.Text
                    Vector(Renglon, 2) = rstOrden!Articulo
                    Vector(Renglon, 3) = Str$(rstOrden!Cantidad)
                    Vector(Renglon, 4) = Str$(rstOrden!Precio)
                    Vector(Renglon, 5) = Str$(rstOrden!Cantidad * rstOrden!Precio)
                    Vector(Renglon, 6) = Str$(rstOrden!Derechos)
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = WCarpeta + "01"
                    
                    WTotalImpo = WTotalImpo + (rstOrden!Cantidad * rstOrden!Precio)
                    WTotalPeso = WTotalPeso + rstOrden!Cantidad
                    
                    WLeyenda = IIf(IsNull(rstOrden!Leyenda), "0", rstOrden!Leyenda)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
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
    
    For Ciclo = 1 To Renglon
        WSeguro = 0
        WFlete = 0
        If ImpoSeguro <> 0 Then
            If WTotalImpo <> 0 Then
                WSeguro = ((Val(Vector(Ciclo, 5)) / WTotalImpo) * ImpoSeguro) / Val(Vector(Ciclo, 3))
            End If
        End If
        If ImpoFlete <> 0 Then
            If WTotalPeso <> 0 Then
                WFlete = ((Val(Vector(Ciclo, 3)) / WTotalPeso) * ImpoFlete) / Val(Vector(Ciclo, 3))
            End If
        End If
        WCosto = Val(Vector(Ciclo, 4)) + WSeguro + WFlete
        Call Redondeo(WCosto)
        Vector(Ciclo, 4) = Str$(WCosto)
    Next Ciclo
    
    WTotal = 0
    
    For Ciclo = 1 To Renglon
        WArticulo = Vector(Ciclo, 2)
        
        spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Rem Vector(Ciclo, 4) = Str$(rstArticulo!Flete)
            Vector(Ciclo, 5) = Str$(Val(Vector(Ciclo, 3) * Val(Vector(Ciclo, 4))))
            XArancel = Val(Vector(Ciclo, 5)) * Val(Vector(Ciclo, 6)) / 100
            Vector(Ciclo, 7) = Str$(Val(Vector(Ciclo, 5)) + XArancel)
            WTotal = WTotal + Val(Vector(Ciclo, 5))
            ImpoArancel = ImpoArancel + XArancel
            rstArticulo.Close
        End If
    Next Ciclo
    
    For Ciclo = 1 To Renglon
        If WTotal <> 0 Then
            Vector(Ciclo, 8) = Str$((Val(Vector(Ciclo, 5)) / WTotal) * ImpoGastos)
        End If
        If Val(Vector(Ciclo, 3)) <> 0 Then
            Vector(Ciclo, 9) = Str$((Val(Vector(Ciclo, 7)) + Val(Vector(Ciclo, 8))) / Val(Vector(Ciclo, 3)))
        End If
    Next Ciclo
    
    If WTotal <> 0 Then
        WCoeficiente = Str$((ImpoGastos + ImpoArancel) / (WTotal / 100))
        WCoeficiente = Str$(1 + (Val(WCoeficiente) / 100))
            Else
        WCoeficiente = ""
    End If
    
    Lugar = 0
    SumaTotal = 0
    SumaGastos = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM MovGasParcialArti"
    Sql3 = " Where MovGasParcialArti.Codigo = " + "'" + Codigo.Text + "'"
    Sql4 = " Order by MovGasParcialArti.Clave"
    spMovGasParcialArti = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcialArti.RecordCount > 0 Then
        With rstMovGasParcialArti
            .MoveFirst
            Do
                If .EOF = False Then
                    Lugar = Lugar + 1
                    Vector(Lugar, 12) = Str$(rstMovGasParcialArti!Cantidad)
                    Vector(Lugar, 13) = Str$(Val(Vector(Lugar, 9)) * Val(Vector(Lugar, 12)))
                    SumaTotal = SumaTotal + Val(Vector(Lugar, 13))
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovGasParcialArti.Close
    End If
    
    
    Sql1 = "Select *"
    Sql2 = " FROM MovGasParcial"
    Sql3 = " Where MovGasParcial.Codigo = " + "'" + Codigo.Text + "'"
    Sql4 = " Order by MovGasParcial.Clave"
    spMovGasParcial = Sql1 + Sql2 + Sql3 + Sql4
    Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovGasParcial.RecordCount > 0 Then
        With rstMovGasParcial
            .MoveFirst
            Do
                If .EOF = False Then
                    SumaGastos = SumaGastos + rstMovGasParcial!Importe
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovGasParcial.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        If SumaTotal <> 0 Then
            Vector(Ciclo, 14) = Str$((Val(Vector(Ciclo, 13)) / SumaTotal) * SumaGastos)
        End If
    
    Next Ciclo
    
    XLeyenda = Str$(WLeyenda)
    
    For Ciclo = 1 To Renglon
    
        ZCarpeta = Vector(Ciclo, 1)
        ZArticulo = Vector(Ciclo, 2)
        ZCantidad = Vector(Ciclo, 3)
        ZCostoFlete = Vector(Ciclo, 4)
        ZImporte = Vector(Ciclo, 5)
        ZArancel = Vector(Ciclo, 6)
        ZCosto = Vector(Ciclo, 7)
        ZGastos = Vector(Ciclo, 8)
        ZPrecio = Vector(Ciclo, 9)
        ZCoeficiente = WCoeficiente
        ZClave = Vector(Ciclo, 11)
        ZLeyenda = XLeyenda
        ZCantidadII = Vector(Ciclo, 12)
        ZGastosII = Vector(Ciclo, 14)
                
        Sql1 = "INSERT INTO Carpeta ("
        Sql2 = "Carpeta ,"
        Sql3 = "Articulo ,"
        Sql4 = "Cantidad ,"
        Sql5 = "CostoFlete ,"
        Sql6 = "Importe ,"
        Sql7 = "Arancel ,"
        Sql8 = "Costo ,"
        Sql9 = "Gastos ,"
        Sql10 = "Precio ,"
        Sql11 = "Coeficiente ,"
        Sql12 = "Clave ,"
        Sql13 = "Leyenda ,"
        Sql14 = "CantidadII ,"
        Sql15 = "GastosII )"
        Sql16 = "Values ("
        Sql17 = "'" + ZCarpeta + "',"
        Sql18 = "'" + ZArticulo + "',"
        Sql19 = "'" + ZCantidad + "',"
        Sql20 = "'" + ZCostoFlete + "',"
        Sql21 = "'" + ZImporte + "',"
        Sql22 = "'" + ZArancel + "',"
        Sql23 = "'" + ZCosto + "',"
        Sql24 = "'" + ZGastos + "',"
        Sql25 = "'" + ZPrecio + "',"
        Sql26 = "'" + ZCoeficiente + "',"
        Sql27 = "'" + ZClave + "',"
        Sql28 = "'" + ZLeyenda + "',"
        Sql29 = "'" + ZCantidadII + "',"
        Sql30 = "'" + ZGastosII + "')"
        spCarpeta = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                    + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 _
                    + Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30
        Set rstCarpeta = db.OpenRecordset(spCarpeta, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Calculo de Costo de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Carpeta.Carpeta} in " + Carpeta.Text + " to " + Carpeta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Carpeta.Carpeta, Carpeta.Articulo, Carpeta.Cantidad, Carpeta.CostoFlete, Carpeta.Importe, Carpeta.Arancel, Carpeta.Costo, Carpeta.Gastos, Carpeta.Precio, Carpeta.Leyenda, Carpeta.CantidadII, Carpeta.GastosII, " _
                    + "Articulo.Descripcion, " _
                    + "Movgas.Fecha, Movgas.Orden, Movgas.Proveedor, Movgas.Origen, Movgas.Moneda, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Carpeta Carpeta, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Movgas Movgas, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Carpeta.Articulo = Articulo.Codigo AND " _
                    + "Carpeta.Clave = Movgas.Clave AND " _
                    + "Movgas.Proveedor = Proveedor.Proveedor AND " _
                    + "Carpeta.Carpeta >= 0 AND " _
                    + "Carpeta.Carpeta <= 999999 AND " _
                    + "Carpeta.CantidadII > 0"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()

    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Carpeta.SetFocus
    PrgCostoOrdenParcial.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Codigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM MovGasParcial"
        Sql3 = " Where MovGasParcial.Codigo = " + "'" + Codigo.Text + "'"
        spMovGasParcial = Sql1 + Sql2 + Sql3
        Set rstMovGasParcial = db.OpenRecordset(spMovGasParcial, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovGasParcial.RecordCount > 0 Then
            Carpeta.Text = rstMovGasParcial!Carpeta
            rstMovGasParcial.Close
                Else
            Sql1 = "Select *"
            Sql2 = " FROM MovGasParcialArti"
            Sql3 = " Where MovGasParcialArti.Codigo = " + "'" + Codigo.Text + "'"
            spMovGasParcialArti = Sql1 + Sql2 + Sql3
            Set rstMovGasParcialArti = db.OpenRecordset(spMovGasParcialArti, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovGasParcialArti.RecordCount > 0 Then
                Carpeta.Text = rstMovGasParcialArti!Carpeta
                rstMovGasParcialArti.Close
            End If
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
    
Sub Form_Load()
    Codigo.Text = ""
    Carpeta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub






