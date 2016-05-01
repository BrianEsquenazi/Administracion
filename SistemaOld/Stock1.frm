VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgStock1 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valorizacion de Materia Prima a Fecha"
   ClientHeight    =   3750
   ClientLeft      =   1050
   ClientTop       =   1305
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   3750
   ScaleWidth      =   9585
   Begin Crystal.CrystalReport Listado 
      Left            =   6240
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wStock1.rpt"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2535
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgStock1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(10000) As String
Dim Empe(10, 10) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
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
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim XParam As String
Dim WFechaord As String
Dim XOrden As Double
Dim XLaudo As Double
Dim XFechaOrden As String
Dim XCostoOrden As Double
Dim XParidad As Double
Dim XMoneda As Integer
Dim XTipoOrden As Integer
Dim WCosto As Double

Private Sub Acepta_Click()

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Posdat = "al " + Fecha.Text
            .Update
        End If
    End With

    Erase Vector
    Renglon = 0

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo

            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstArticulo!Codigo >= Desde.Text And rstArticulo!Codigo <= Hasta.Text Then
                    Renglon = Renglon + 1
                    Vector(Renglon) = rstArticulo!Codigo
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    rstArticulo.Close
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaord = WAno + WMes + WDia
    
    spArticulo = "ModificaArticuloStock0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WArticulo = Vector(Da)
        XCodigo = Vector(Da)
        XDate = Date$
        
        Call calcula_datos
        
        WStock = Str$(WEntradas - WSalidas)
        
        XParam = "'" + XCodigo + "','" _
                + WStock + "'"
                                           
        spArticulo = "ModificaArticuloStock " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem actualiza el costo
        
        spArticulo = "ConsultaArticulo " + "'" + XCodigo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = rstArticulo!Costo2
            rstArticulo.Close
        End If
        
        XOrden = 0
        XLaudo = 0
        XFechaOrden = "00000000"
        XCostoOrden = 0
        XMoneda = 0
        XTipoOrden = 0
        
        XEmpresa = WEmpresa
    
        If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 9 Or Val(XEmpresa) = 10 Then
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
            XHasta = 5
                Else
            Empe(1, 1) = "0002"
            Empe(1, 2) = "Empresa02"
            Empe(2, 1) = "0004"
            Empe(2, 2) = "Empresa04"
            Empe(3, 1) = "0008"
            Empe(3, 2) = "Empresa08"
            XHasta = 3
        End If
    
        For A = 1 To XHasta
        
            WEmpresa = Empe(A, 1)
            txtOdbc = Empe(A, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            XOrden = 0
            
            XParam = "'" + XCodigo + "','" _
                 + XCodigo + "'"
            spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                With rstLaudo
                    .MoveFirst
                    If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                        If rstLaudo!Articulo = XCodigo Then
                            XOrdFecha = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            If XOrdFecha > XFechaOrden Then
                                XFechaOrden = XOrdFecha
                                XOrden = !Orden
                                XLaudo = !Laudo
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
            
            If XOrden <> 0 Then
                spOrden = "ListaOrdenArticulo " + "'" + Str$(XOrden) + "','" + XCodigo + "'"
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                If rstOrden.RecordCount > 0 Then
                    XTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
                    If XTipoOrden = 0 And rstOrden!Precio <> 1 Then
                        XCostoOrden = rstOrden!Precio
                        XMoneda = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
                            Else
                        vercodigo = XCodigo
                        XFechaOrden = WFechaord
                        XCostoOrden = WCosto
                        XMoneda = 0
                    End If
                    rstOrden.Close
                End If
                If XMoneda = 0 Then
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    spCambios = "ConsultaCambioOrdFecha  " + "'" + XFechaOrden + "'"
                    Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCambios.RecordCount > 0 Then
                        With rstCambios
                            .MoveLast
                            XParidad = rstCambios!Cambio
                            rstCambios.Close
                        End With
                            Else
                        XParidad = 1
                    End If
                    
                    WEmpresa = Empe(A, 1)
                    txtOdbc = Empe(A, 2)
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    XCostoOrden = XCostoOrden * XParidad
                End If
            End If
            
        Next A
        
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
            Case Else
        End Select
        
        If XCostoOrden = 0 Then
            XCostoOrden = WCosto
        End If
        
        XParam = "'" + XCodigo + "','" _
                    + Str$(XCostoOrden) + "'"
        spArticulo = "ModificaArticuloCostoImpre " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da
    
    Listado.WindowTitle = "Listado de Valorizacion de Materia Prima a Fecha"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaEnv.Envase} in " + DesdeEnv.Text + " to " + HastaEnv.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Costo2, Articulo.Stock " _
                        + "From " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where " _
                        + "Articulo.Codigo >= '  -   -   ' AND " _
                        + "Articulo.Codigo <= 'ZZ-999-999' AND Articulo.Stock <> 0."
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1

End Sub

Private Sub calcula_datos()

    Rem PROCESA LOS LAUDOS
    
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
        rstArticulo.Close
    End If
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    spLaudo = "ListaLaudoRepro" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)

                If rstLaudo!Marca = "X" And WSaldo = 0 Then
                
                        Else
                        
                    WAno = Right$(rstLaudo!Fecha, 4)
                    WMes = Mid$(rstLaudo!Fecha, 4, 2)
                    WDia = Left$(rstLaudo!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                        
                    If WCompara <= WFechaord Then
                        If rstLaudo!Articulo = WArticulo Then
                            WEntradas = WEntradas + rstLaudo!Liberada
                        End If
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
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    Rem spHoja = "ListaHojaRepro" + XParam
    spHoja = "ListaHojaArticuloDesdeHasta" + XParam
    Rem spHoja = "ListaHojaRepro " + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                        Else
                        
                    WAno = Right$(rstHoja!Fecha, 4)
                    WMes = Mid$(rstHoja!Fecha, 4, 2)
                    WDia = Left$(rstHoja!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                        
                    If WCompara <= WFechaord Then
                        If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                            XX = rstHoja!Clave
                            WSalidas = WSalidas + rstHoja!Cantidad
                        End If
                    End If
                    
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstHoja!Articulo > WArticulo Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
        
        rstHoja.Close
        
    End If
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovvar = "ListaMovvarRepro1" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                        
                    WAno = Right$(rstMovvar!Fecha, 4)
                    WMes = Mid$(rstMovvar!Fecha, 4, 2)
                    WDia = Left$(rstMovvar!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                        
                    If WCompara <= WFechaord Then
                        
                        If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                            If rstMovvar!Movi = "E" Then
                                WEntradas = WEntradas + rstMovvar!Cantidad
                                    Else
                                WSalidas = WSalidas + rstMovvar!Cantidad
                            End If
                        End If
                        
                    End If
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
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                + WArticulo + "'"
    spMovguia = "ListaMovguiaRepro1" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                        
                    WAno = Right$(rstMovguia!Fecha, 4)
                    WMes = Mid$(rstMovguia!Fecha, 4, 2)
                    WDia = Left$(rstMovguia!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                        
                    If WCompara <= WFechaord Then
                        
                        If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                            If rstMovguia!Movi = "E" Then
                                WEntradas = WEntradas + rstMovguia!Cantidad
                                    Else
                                WSalidas = WSalidas + rstMovguia!Cantidad
                            End If
                        End If
                            
                    End If
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
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spMovlab = "ListaMovlabRepro1" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovlab!Marca = "X" Then
                
                        Else
                        
                    WAno = Right$(rstMovlab!Fecha, 4)
                    WMes = Mid$(rstMovlab!Fecha, 4, 2)
                    WDia = Left$(rstMovlab!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                        
                    If WCompara <= WFechaord Then
                
                        If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                            If rstMovlab!Movi = "E" Then
                                WEntradas = WEntradas + rstMovlab!Cantidad
                                    Else
                                WSalidas = WSalidas + rstMovlab!Cantidad
                            End If
                        End If
                        
                    End If
                    
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    
    Rem PROCESA LAS VENTAS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHasta" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEstadistica!Marca = "X" Then
                
                        Else
                
                    If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
                    
                        WAno = Right$(rstEstadistica!Fecha, 4)
                        WMes = Mid$(rstEstadistica!Fecha, 4, 2)
                        WDia = Left$(rstEstadistica!Fecha, 2)
                        WCompara = WAno + WMes + WDia
                        
                        If WCompara <= WFechaord Then
                            If rstEstadistica!Tipo = 1 Then
                                WSalidas = WSalidas + rstEstadistica!Cantidad
                                    Else
                                WEntradas = WEntradas + rstEstadistica!Cantidad
                            End If
                        End If
                
                    End If
                    
                End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Fecha.SetFocus
    PrgStock1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Fecha.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgStock1.Caption = "Listado de Valorizacion de Materia Prima a Fecha :  " + !Nombre
        End If
    End With
    
    Fecha.Text = "  /  /    "
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

