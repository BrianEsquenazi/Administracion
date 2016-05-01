VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMoreno 
   AutoRedraw      =   -1  'True
   Caption         =   "2.-Listado de Estadisitica de Ventas por Rubro y Clientes"
   ClientHeight    =   6525
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   8145
   Begin VB.CommandButton Command3 
      Caption         =   "Exportaciones"
      Height          =   1335
      Left            =   3120
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Importaciones"
      Height          =   1335
      Left            =   4680
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lista de Precios"
      Height          =   1335
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WEsta1.rpt"
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
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgMoreno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Costo As Double
Private Producto As String
Private Auxiliar(100, 7) As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstLinea As Recordset
Dim spLinea As String
Dim rstRubro As Recordset
Dim spRubro As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Private Vecosto(5000, 2) As String
Dim Posi As Integer
Dim ZDescriLinea(100) As String
Dim ZDescriRubro(100) As String
Dim ZZImpo As Double







Private Sub Command1_Click()

    Rem On Error GoTo WError
    
    OPEN_FILE_Moreno
    
    With rstMoreno
        .Index = "Articulo"
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
    
    WDesde = "20090201"
    Whasta = "20120131"
    
    WDesdeI = "20110201"
    WHastaI = "20120131"
    
    WDesdeII = "20100201"
    WHastaII = "20110131"
    
    WDesdeIII = "20090201"
    WHastaIII = "20100131"
    
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + Whasta + "'"
    ZSql = ZSql + " Order by Articulo"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
            
            
                If rstEstadistica!numero <= 800000 Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstEstadistica!Articulo
                        ZPrecioI = 0
                        ZPrecioII = 0
                        ZPrecioIII = 0
                        ZClienteI = ""
                        ZClienteII = ""
                        ZClienteIII = ""
                        ZZPrecioI = 0
                        ZZPrecioII = 0
                        ZZPrecioIII = 0
                        ZZClienteI = ""
                        ZZClienteII = ""
                        ZZClienteIII = ""
                    End If
                    
                    If ZCorte <> rstEstadistica!Articulo Then
                    
                        With rstMoreno
                            .Index = "Articulo"
                            .AddNew
                            !Articulo = ZCorte
                            !Descripcion = IIf(WDescriArticulo = "", "0", WDescriArticulo)
                            !precioI = ZZPrecioI
                            !ClienteI = ZZClienteI
                            !precioII = ZZPrecioII
                            !ClienteII = ZZClienteII
                            !precioIII = ZZPrecioIII
                            !ClienteIII = ZZClienteIII
                            .Update
                        End With
                    
                        ZCorte = rstEstadistica!Articulo
                        ZPrecioI = 0
                        ZPrecioII = 0
                        ZPrecioIII = 0
                        ZClienteI = ""
                        ZClienteII = ""
                        ZClienteIII = ""
                        ZZPrecioI = 0
                        ZZPrecioII = 0
                        ZZPrecioIII = 0
                        ZZClienteI = ""
                        ZZClienteII = ""
                        ZZClienteIII = ""

                    End If
                       
                    ZZImpo = rstEstadistica!PrecioUs
                    Call Redondeo(ZZImpo)
                       
                    If rstEstadistica!OrdFecha >= WDesdeI And rstEstadistica!OrdFecha <= WHastaI Then
                        Rem 1
                        If rstEstadistica!Articulo = "PT-30126-020" Then Stop
                        If ZZImpo >= ZPrecioI Then
                            If ZPrecioI > ZZPrecioI And ZZImpo > ZPrecioI Then
                                ZZPrecioI = ZPrecioI
                                ZZClienteI = ZClienteI
                            End If
                            If ZZPrecioI = 0 Then
                                ZZPrecioI = ZPrecioI
                                ZZClienteI = ZClienteI
                            End If
                            ZPrecioI = ZZImpo
                            ZClienteI = rstEstadistica!Cliente
                                Else
                            If ZZImpo > ZZPrecioI Then
                                ZZPrecioI = ZZImpo
                                ZZClienteI = rstEstadistica!Cliente
                            End If
                        End If
                    End If
                       
                    If rstEstadistica!OrdFecha >= WDesdeII And rstEstadistica!OrdFecha <= WHastaII Then
                        Rem 2
                        If rstEstadistica!Articulo = "PT-30126-020" Then Stop
                        If ZZImpo >= ZPrecioII Then
                            If ZPrecioII > ZZPrecioIIAnd And ZZImpo > ZPrecioII Then
                                ZZPrecioII = ZPrecioII
                                ZZClienteII = ZClienteII
                            End If
                            If ZZPrecioII = 0 Then
                                ZZPrecioII = ZPrecioII
                                ZZClienteII = ZClienteII
                            End If
                            ZPrecioII = ZZImpo
                            ZClienteII = rstEstadistica!Cliente
                                Else
                            If ZZImpo > ZZPrecioII Then
                                ZZPrecioII = ZZImpo
                                ZZClienteII = rstEstadistica!Cliente
                            End If
                        End If
                    End If
                       
                    If rstEstadistica!OrdFecha >= WDesdeIII And rstEstadistica!OrdFecha <= WHastaIII Then
                        Rem 3
                        If rstEstadistica!Articulo = "PT-30126-020" Then Stop
                        If ZZImpo >= ZPrecioIII Then
                            If ZPrecioIII > ZZPrecioIII And ZZImpo > ZPrecioIII Then
                                ZZPrecioIII = ZPrecioIII
                                ZZClienteIII = ZClienteIII
                            End If
                            If ZZPrecioIII = 0 Then
                                ZZPrecioIII = ZPrecioIII
                                ZZClienteIII = ZClienteIII
                            End If
                            ZPrecioIII = ZZImpo
                            ZClienteIII = rstEstadistica!Cliente
                                Else
                            If ZZImpo > ZZPrecioIII Then
                                ZZPrecioIII = ZZImpo
                                ZZClienteIII = rstEstadistica!Cliente
                            End If
                        End If
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    
    
    With rstMoreno
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WDescriArticulo = ""
                        
               If Left$(!Articulo, 2) = "PT" Or Left$(!Articulo, 2) = "PE" Then
                    
                    spTerminado = "ConsultaTerminado" + "'" + !Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WDescriArticulo = IIf(rstTerminado!Descripcion = "", "0", rstTerminado!Descripcion)
                        rstTerminado.Close
                    End If
                    
                        Else
                        
                    XArti = Left$(!Articulo, 3) + Right$(!Articulo, 7)
                    spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescriArticulo = IIf(rstArticulo!Descripcion = "", "0", rstArticulo!Descripcion = "")
                        rstArticulo.Close
                    End If
                    
                End If
                    
                !Descripcion = IIf(IsNull(WDescriArticulo), "0", Wdescrarticulo)
                    
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
                        
    
    
    
    Stop
    
    
    Listado.WindowTitle = "2.-Listado de Estadistica de Ventas por Rubro y Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    Desde = "1"
    Hasta = "20"
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Dos = " and {Estadistica.Rubro} in " + Desde + " to " + Hasta
    Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WEsta1ii.rpt"
     
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEsta
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgEsta1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Private Sub DesdeFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Private Sub Command2_Click()


    Rem On Error GoTo WError
    
    Dim ZVector(15000, 10) As String
    Dim DescriImpo(100) As String
    
    DescriImpo(1) = "FOB"
    DescriImpo(2) = "CIF"
    DescriImpo(3) = "CFR"
    DescriImpo(4) = "CPT"
    DescriImpo(5) = "EXW"
    DescriImpo(6) = "FCA"
    
    
    OPEN_FILE_Moreno
    
    
    With rstMoreno
        .Index = "Articulo"
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
    
    
    WDesde = "20090201"
    Whasta = "20120131"
    
    WDesdeI = "20110201"
    WHastaI = "20120131"
    
    WDesdeII = "20100201"
    WHastaII = "20110131"
    
    WDesdeIII = "20090201"
    WHastaIII = "20100131"
    
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Laudo.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Laudo.FechaOrd <= " + "'" + Whasta + "'"
    ZSql = ZSql + " Order by Articulo"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            Do
            
                WLiberada = rstLaudo!Liberada
                WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                If WLiberadaAnt > 0 Then
                    ZLiberada = WLiberadaAnt
                        Else
                    ZLiberada = WLiberada
                End If
            
                If rstLaudo!FechaOrd >= WDesdeI And rstLaudo!FechaOrd <= WHastaI Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstLaudo!Laudo
                    ZVector(ZLugar, 2) = rstLaudo!Articulo
                    ZVector(ZLugar, 3) = Str$(ZLiberada)
                    ZVector(ZLugar, 4) = ""
                    ZVector(ZLugar, 5) = ""
                    ZVector(ZLugar, 6) = rstLaudo!Orden
                    ZVector(ZLugar, 7) = rstLaudo!Fecha
                End If
                   
                If rstLaudo!FechaOrd >= WDesdeII And rstLaudo!FechaOrd <= WHastaII Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstLaudo!Laudo
                    ZVector(ZLugar, 2) = rstLaudo!Articulo
                    ZVector(ZLugar, 3) = ""
                    ZVector(ZLugar, 4) = Str$(ZLiberada)
                    ZVector(ZLugar, 5) = ""
                    ZVector(ZLugar, 6) = rstLaudo!Orden
                    ZVector(ZLugar, 7) = rstLaudo!Fecha
                End If
                   
                If rstLaudo!FechaOrd >= WDesdeIII And rstLaudo!FechaOrd <= WHastaIII Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstLaudo!Laudo
                    ZVector(ZLugar, 2) = rstLaudo!Articulo
                    ZVector(ZLugar, 3) = ""
                    ZVector(ZLugar, 4) = ""
                    ZVector(ZLugar, 5) = Str$(ZLiberada)
                    ZVector(ZLugar, 6) = rstLaudo!Orden
                    ZVector(ZLugar, 7) = rstLaudo!Fecha
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZLaudo = ZVector(Ciclo, 1)
        ZArticulo = ZVector(Ciclo, 2)
        ZCantiI = Val(ZVector(Ciclo, 3))
        ZCantiII = Val(ZVector(Ciclo, 4))
        ZCantiIII = Val(ZVector(Ciclo, 5))
        ZOrden = ZVector(Ciclo, 6)
        ZFecha = ZVector(Ciclo, 7)
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Orden = " + "'" + ZOrden + "'"
        ZSql = ZSql + " and Orden.Articulo = " + "'" + ZArticulo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            If rstOrden!Tipo = 1 Then
            
                ZPrecio = rstOrden!Precio
                ZZZPrecio = rstOrden!Precio
                ZLeyenda = rstOrden!leyenda
                ZMoneda = rstOrden!moneda
                ZFecha = rstOrden!Fecha
                rstOrden.Close
                
                If ZMoneda = 2 Then
                
                    XEmpresa = WEmpresa
                    
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    ZZFactor = 1.35
                    spCambios = "ConsultaCambio " + "'" + ZFecha + "'"
                    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCambios.RecordCount > 0 Then
                        ZZDolar = rstCambios!Cambio
                        ZCambioII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
                        ZZEuro = ZCambioII
                        If ZZDolar <> 0 And ZZEuro <> 0 Then
                            ZZFactor = ZZEuro / ZZDolar
                        End If
                        rstCambios.Close
                    End If
                    ZPrecio = ZPrecio * ZZFactor
                    
                    Call Conecta_Empresa
                    
                End If
                
                If ZLeyenda <> 1 Then
                    Rem ZPrecio = ZPrecio / 1.03
                End If

                ZPrecioI = ZCantiI * ZPrecio
                ZZZPrecioI = ZCantiI * ZZZPrecio
                ZPrecioII = ZCantiII * ZPrecio
                ZPrecioIII = ZCantiIII * ZPrecio
                
                Dife = ZZZPrecioI - ZPrecioI
                Suma = Suma + Dife
                
                
                WDescriArticulo = ""
    
                With rstMoreno
                    .Index = "Articulo"
                    .AddNew
                    !Articulo = ZArticulo
                    !Descripcion = WDescriArticulo
                    !precioI = ZPrecioI
                    If ZLeyenda < 1 Then
                        ZLeyenda = 0
                    End If
                    !ClienteI = DescriImpo(ZLeyenda)
                    !precioII = ZPrecioII
                    !ClienteII = ZLaudo
                    !precioIII = ZPrecioIII
                    !ClienteIII = ZOrden
                    .Update
                End With
                
            End If
            
        End If
        
    Next Ciclo
    
                        
    
    With rstMoreno
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WDescriArticulo = ""
                
                ZArticulo = !Articulo
                        
                spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                    
                !Descripcion = WDescriArticulo
                    
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    aaa = Suma
    
    
    
    
    Stop
    
    
    Listado.WindowTitle = "2.-Listado de Estadistica de Ventas por Rubro y Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    Desde = "1"
    Hasta = "20"
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Dos = " and {Estadistica.Rubro} in " + Desde + " to " + Hasta
    Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WEsta1ii.rpt"
     
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    


End Sub

Private Sub Command3_Click()

    Rem On Error GoTo WError
    
    OPEN_FILE_Moreno
    
    With rstMoreno
        .Index = "Articulo"
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
    
    WDesde = "20090201"
    Whasta = "20120101"
    
    WDesdeI = "20111201"
    WHastaI = "20111231"
    
    WDesdeII = "20100201"
    WHastaII = "20110131"
    
    WDesdeIII = "20090201"
    WHastaIII = "20100131"
    
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + Whasta + "'"
    ZSql = ZSql + " Order by Articulo"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            Do
            
            
                If rstEstadistica!numero > 800000 Then
                If rstEstadistica!Cliente <> "A00066" Then
                
                    If Val(rstEstadistica!Tipo) <> 1 Then Stop
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstEstadistica!Articulo
                        ZPrecioI = 0
                        ZPrecioII = 0
                        ZPrecioIII = 0
                    End If
                    
                    If ZCorte <> rstEstadistica!Articulo Then
                    
                        With rstMoreno
                            .Index = "Articulo"
                            .AddNew
                            !Articulo = ZCorte
                            !Descripcion = IIf(WDescriArticulo = "", "0", WDescriArticulo)
                            !precioI = ZPrecioI
                            !precioII = ZPrecioII
                            !precioIII = ZPrecioIII
                            .Update
                        End With
                    
                        ZCorte = rstEstadistica!Articulo
                        ZPrecioI = 0
                        ZPrecioII = 0
                        ZPrecioIII = 0

                    End If
                       
                    ZZImpo = rstEstadistica!Precio * rstEstadistica!Cantidad
                    Call Redondeo(ZZImpo)
                       
                    If rstEstadistica!OrdFecha >= WDesdeI And rstEstadistica!OrdFecha <= WHastaI Then
                        ZPrecioI = ZPrecioI + ZZImpo
                    End If
                       
                    If rstEstadistica!OrdFecha >= WDesdeII And rstEstadistica!OrdFecha <= WHastaII Then
                        ZPrecioII = ZPrecioII + ZZImpo
                    End If
                       
                    If rstEstadistica!OrdFecha >= WDesdeIII And rstEstadistica!OrdFecha <= WHastaIII Then
                        ZPrecioIII = ZPrecioIII + ZZImpo
                    End If
                    
                End If
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End With
    End If
    
    
    
    With rstMoreno
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WDescriArticulo = ""
                        
               If Left$(!Articulo, 2) = "PT" Or Left$(!Articulo, 2) = "PE" Then
                    
                    spTerminado = "ConsultaTerminado" + "'" + !Articulo + "'"
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    If rstTerminado.RecordCount > 0 Then
                        WDescriArticulo = IIf(rstTerminado!Descripcion = "", "0", rstTerminado!Descripcion)
                        rstTerminado.Close
                    End If
                    
                        Else
                        
                    XArti = Left$(!Articulo, 3) + Right$(!Articulo, 7)
                    spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WDescriArticulo = IIf(rstArticulo!Descripcion = "", "0", rstArticulo!Descripcion = "")
                        rstArticulo.Close
                    End If
                    
                End If
                    
                !Descripcion = IIf(IsNull(WDescriArticulo), "0", Wdescrarticulo)
                    
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
                        
    
    
    
    Stop
    
    
    Listado.WindowTitle = "2.-Listado de Estadistica de Ventas por Rubro y Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    Desde = "1"
    Hasta = "20"
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    Dos = " and {Estadistica.Rubro} in " + Desde + " to " + Hasta
    Listado.GroupSelectionFormula = Uno + Dos
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    Listado.ReportFileName = "WEsta1ii.rpt"
     
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    

End Sub

Sub Form_Load()
End Sub


Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    If Left$(Producto, 2) = "PT" Or Left$(Producto, 2) = "PE" Or Left$(Producto, 2) = "DW" Or Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
    
    If Left$(Producto, 2) = "NK" Or Left$(Producto, 2) = "RE" Then
        Producto = "PT" + Mid$(Producto, 3, 10)
    End If
    
    If Left$(Producto, 2) = "NW" Then
        Producto = "DW" + Mid$(Producto, 3, 10)
    End If

    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    For da = 1 To Posi
        If Producto = Vecosto(da, 1) Then
            Costo = Val(Vecosto(da, 2))
            Exit Sub
        End If
    Next da
    
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
                        
                        If Left$(Articulo1, 2) = "DW" Then
                            Tipo = "T"
                            Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        End If
                        
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
            
            If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
                Renglon = Renglon + 1
                Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
                Auxiliar(Renglon, 2) = 1
                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
    
    If Renglon > 0 Then
                    
        For da = 1 To Renglon
            Articulo = Auxiliar(da, 1)
            Cantidad = Auxiliar(da, 2)
            XVector = Auxiliar(da, 3)
            
            spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
                Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
                rstArticulo.Close
            End If
        Next da
        
            Else
            
        XArti = Left$(Producto, 3) + Right$(Producto, 7)
        spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Costo = rstArticulo!Costo2
            rstArticulo.Close
        End If
    
    End If
            
    
    Posi = Posi + 1
    Vecosto(Posi, 1) = Producto
    Vecosto(Posi, 2) = Str$(Costo)
    
        Else
        
    XArti = Left$(Producto, 3) + Right$(Producto, 7)
    spArticulo = "ConsultaArticulo " + "'" + XArti + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Costo = rstArticulo!Costo2
        rstArticulo.Close
    End If

    End If
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Esta
End Sub

