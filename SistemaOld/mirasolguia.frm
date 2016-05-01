VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgMiraSolGuia 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Solicitud de Emision de Guias de Traslado Interno"
   ClientHeight    =   7320
   ClientLeft      =   90
   ClientTop       =   585
   ClientWidth     =   11760
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   11760
   Begin VB.CommandButton Anula 
      Caption         =   "Anula Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Guia 
      Caption         =   "Guia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Impresion 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   4000
      Cols            =   10
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wsolguiatotal.rpt"
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Lee datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "PrgMiraSolGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstSolGuia As Recordset
Dim spSolGuia As String
Dim rstSolGuiaTotal As Recordset
Dim spSolGuiaTotal As String
Dim XParam As String
Dim WGraba As String
Private TotalSolicitud As Integer
Dim Auxiliar(10000, 15) As String
Dim ZSolicitud As String
Dim ZFecha As String
Dim ZArticulo As String
Dim ZTerminado As String
Dim ZCantidad As String
Dim ZDesde As String
Dim ZHasta As String
Dim ZObservaciones As String
Dim ZEmpresa As String
Dim ZClave As String
Dim ZTipo As String
Dim ZOrdFecha As String
Dim ZFechaOrd As String
Dim XProceso As Integer


Private Sub Anula_Click()

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

    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    For Ciclo = RowIni To Rowfin
    
        ClaveSol = Auxiliar(Ciclo, 1)
        WMarca = "X"
        XParam = "'" + ClaveSol + "','" _
                     + WMarca + "'"

        spSolGuia = "ModificaSolGuiaMarca " + XParam
        Set rstSolGuia = db.OpenRecordset(spSolGuia, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Call Conecta_Empresa
    
    Call Proceso_Click

End Sub

Private Sub cmdClose_Click()
    With rstEmpresa
        .Close
    End With
    PrgMiraSolGuia.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    XEmpresa = WEmpresa

    Call Limpia_Vector
    
    Muestra.ColWidth(0) = 50
    Muestra.ColWidth(1) = 850
    Muestra.ColWidth(2) = 1200
    Muestra.ColWidth(3) = 1400
    Muestra.ColWidth(4) = 2200
    Muestra.ColWidth(5) = 900
    Muestra.ColWidth(6) = 800
    Muestra.ColWidth(7) = 800
    Muestra.ColWidth(8) = 2400
    Muestra.ColWidth(9) = 700

    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Producto"
    
    Muestra.Col = 4
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 5
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 6
    Muestra.Text = "Origen"
    
    Muestra.Col = 7
    Muestra.Text = "Destino"
    
    Muestra.Col = 8
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 9
    Muestra.Text = "Resp."
    
    XProceso = 0
    
    Call Proceso_Click
    
End Sub

Private Sub GuiaXX_Click()

    RowIni = Muestra.Row
    Rowfin = Muestra.RowSel
    
    Control = Muestra.TextMatrix(Muestra.Row, 6)
    Control1 = Muestra.TextMatrix(Muestra.Row, 7)
    
    PasaDatos = "S"
    
    For Ciclo = RowIni To Rowfin
        If Muestra.TextMatrix(Ciclo, 6) <> Control Then
            PasaDatos = "N"
        End If
    Next Ciclo
    
    If PasaDatos = "N" Then
        m$ = "El Origen de los productos seleccionados deben ser iguales"
        G% = MsgBox(m$, 0, "Guias de Traslado Internos")
        Exit Sub
    End If
    
    PasaDatos = "S"
    
    For Ciclo = RowIni To Rowfin
        If Muestra.TextMatrix(Ciclo, 7) <> Control1 Then
            PasaDatos = "N"
        End If
    Next Ciclo
    
    If PasaDatos = "N" Then
        m$ = "El Destino de los productos seleccionados deben ser iguales"
        G% = MsgBox(m$, 0, "Guias de Traslado Internos")
        Exit Sub
    End If
    
    Erase TraspaDatos
    LugarTraspa = 0
    
    For Ciclo = RowIni To Rowfin
    
        LugarTraspa = LugarTraspa + 1
    
        TraspaDatos(LugarTraspa, 1) = Str$(Val(WEmpresa))
        TraspaDatos(LugarTraspa, 2) = Auxiliar(Ciclo, 8)
        TraspaDatos(LugarTraspa, 3) = Auxiliar(Ciclo, 4)
        TraspaDatos(LugarTraspa, 4) = Auxiliar(Ciclo, 7)
        TraspaDatos(LugarTraspa, 5) = Auxiliar(Ciclo, 5)
        TraspaDatos(LugarTraspa, 6) = Auxiliar(Ciclo, 8)
        TraspaDatos(LugarTraspa, 7) = Auxiliar(Ciclo, 1)
        TraspaDatos(LugarTraspa, 8) = Auxiliar(Ciclo, 6)
        
    Next Ciclo
    
    PasaEmpresa = WEmpresa
    Select Case Control
        Case "SI"
            XEmpresa = "1"
        Case "PI"
            XEmpresa = "2"
        Case "SII"
            XEmpresa = "3"
        Case "PII"
            XEmpresa = "4"
        Case "SIII"
            XEmpresa = "5"
        Case "SIV"
            XEmpresa = "6"
        Case "SV"
            XEmpresa = "7"
        Case "PV"
            XEmpresa = "8"
        Case "PVI"
            XEmpresa = "9"
        Case "SVI"
            XEmpresa = "10"
        Case "SVII"
            XEmpresa = "11"
        Case Else
    End Select
    
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
    
    PrgMiraSolGuia.Hide
    Unload Me
    PrgMovguiaAuto.Show

End Sub

Private Sub Impresion_Click()

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
    
    Listado.WindowTitle = "Listado de Solicitudes de Guia de Traslado Interno"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Terminado.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.Destination = 1
    Listado.Destination = 0
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Call Conecta_Empresa
    
    Listado.SQLQuery = "SELECT SolGuiaTotal.Codigo, SolGuiaTotal.Fecha, SolGuiaTotal.Tipo, SolGuiaTotal.Articulo, SolGuiaTotal.Terminado, SolGuiaTotal.Cantidad, SolGuiaTotal.Desde, SolGuiaTotal.Hasta, SolGuiaTotal.Observaciones, SolGuiaTotal.DescriArticulo, SolGuiaTotal.DescriTerminado " _
                        + "From " _
                        + DSQ + ".dbo.SolGuiaTotal SolGuiaTotal " _
                        + "Where " _
                        + "SolGuiaTotal.Codigo >= 0 AND SolGuiaTotal.Codigo <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Proceso_Click()
    
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
    
    spSolGuiaTotal = "BorrarSolGuiaTotal "
    Set rstSolGuiaTotal = db.OpenRecordset(spSolGuiaTotal, dbOpenSnapshot, dbSQLPassThrough)

    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Producto"
    
    Muestra.Col = 4
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 5
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 6
    Muestra.Text = "Origen"
    
    Muestra.Col = 7
    Muestra.Text = "Destino"
    
    Muestra.Col = 8
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 9
    Muestra.Text = "Resp."
    
    Erase Auxiliar
    WLugar = 0
        
    XParam = "'" + "N" + " '"
    spSolGuia = "ListaSolGuiaPendiente " + XParam
    Set rstSolGuia = db.OpenRecordset(spSolGuia, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolGuia.RecordCount > 0 Then
            
        With rstSolGuia
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Pasa = "N"
                    Select Case Val(XEmpresa)
                        Case 1
                            If Val(rstSolGuia!Desde) = 1 Then
                                Pasa = "S"
                            End If
                        Case 3
                            If Val(rstSolGuia!Desde) = 2 Then
                                Pasa = "S"
                            End If
                        Case 5
                            If Val(rstSolGuia!Desde) = 3 Then
                                Pasa = "S"
                            End If
                        Case 6
                            If Val(rstSolGuia!Desde) = 4 Then
                                Pasa = "S"
                            End If
                        Case 7
                            If Val(rstSolGuia!Desde) = 5 Then
                                Pasa = "S"
                            End If
                        Case Else
                    End Select
                            
                    If Pasa = "S" Then
                        
                        WLugar = WLugar + 1
                        Auxiliar(WLugar, 1) = Pusing("######", Str$(rstSolGuia!Codigo))
                        Auxiliar(WLugar, 2) = rstSolGuia!Fecha
                        Auxiliar(WLugar, 3) = rstSolGuia!Articulo
                        Auxiliar(WLugar, 4) = rstSolGuia!Terminado
                        Auxiliar(WLugar, 5) = Str$(rstSolGuia!Cantidad)
                        Auxiliar(WLugar, 6) = rstSolGuia!Hasta
                        Auxiliar(WLugar, 7) = rstSolGuia!Observaciones
                        Auxiliar(WLugar, 8) = WEmpresa
                        Auxiliar(WLugar, 9) = rstSolGuia!Clave
                        Auxiliar(WLugar, 10) = rstSolGuia!Tipo
                        Auxiliar(WLugar, 11) = rstSolGuia!Desde
                        Auxiliar(WLugar, 12) = rstSolGuia!Usuario
                        
                    End If
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        
        rstSolGuia.Close
        
    End If
    
    For Cicla = 1 To WLugar
    
        ZSolicitud = Auxiliar(Cicla, 1)
        ZFecha = Auxiliar(Cicla, 2)
        ZArticulo = Auxiliar(Cicla, 3)
        ZTerminado = Auxiliar(Cicla, 4)
        ZCantidad = Auxiliar(Cicla, 5)
        ZHasta = Auxiliar(Cicla, 6)
        ZObservaciones = Auxiliar(Cicla, 7)
        ZEmpresa = Auxiliar(Cicla, 8)
        ZClave = Auxiliar(Cicla, 9)
        ZTipo = Auxiliar(Cicla, 10)
        ZDesde = Auxiliar(Cicla, 11)
        ZUsuario = Auxiliar(Cicla, 12)
        ZOrdFecha = "00000000"
        ZFechaOrd = "0000000"
        ZDescriArticulo = ""
        ZDescriTerminado = ""
        
        spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescriArticulo = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        spTerminado = "ConsultaTerminado " + "'" + ZTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZDescriTerminado = rstTerminado!Descripcion
            rstTerminado.Close
        End If
        
        XParam = "'" + ZEmpresa + "','" _
                + ZClave + "','" _
                + ZSolicitud + "','" _
                + ZFecha + "','" _
                + ZFechaOrd + "','" _
                + ZTipo + "','" _
                + ZArticulo + "','" _
                + ZTerminado + "','" _
                + ZCantidad + "','" _
                + ZDesde + "','" _
                + ZHasta + "','" _
                + ZObservaciones + "','" _
                + ZUsuario + "','" _
                + ZDescriArticulo + "','" _
                + ZDescriTerminado + "'"
                         
        spSolGuiaTotal = "AltaSolGuiaTotal " + XParam
        Set rstSolGuiaTotal = db.OpenRecordset(spSolGuiaTotal, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Cicla
    
    Renglon = 0
    Erase Auxiliar
    
    spSolGuiaTotal = "ListaSolGuiaTotal "
    Set rstSolGuiaTotal = db.OpenRecordset(spSolGuiaTotal, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolGuiaTotal.RecordCount > 0 Then
        With rstSolGuiaTotal
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Renglon = Renglon + 1
                    Muestra.Row = Renglon
                                                
                    Muestra.Col = 1
                    Muestra.Text = Pusing("######", Str$(rstSolGuiaTotal!Codigo))
                                            
                    Muestra.Col = 2
                    Muestra.Text = rstSolGuiaTotal!Fecha
                    
                    Select Case rstSolGuiaTotal!Tipo
                        Case "M"
                            Muestra.Col = 3
                            Muestra.Text = rstSolGuiaTotal!Articulo
                        Case Else
                            Muestra.Col = 3
                            Muestra.Text = rstSolGuiaTotal!Terminado
                    End Select
                        
                    Muestra.Col = 5
                    Muestra.Text = rstSolGuiaTotal!Cantidad
                        
                    Select Case Val(rstSolGuiaTotal!Desde)
                        Case 1
                            Muestra.Col = 6
                            Muestra.Text = "SI"
                        Case 2
                            Muestra.Col = 6
                            Muestra.Text = "SII"
                        Case 3
                            Muestra.Col = 6
                            Muestra.Text = "SIII"
                        Case 4
                            Muestra.Col = 6
                            Muestra.Text = "SIV"
                        Case 5
                            Muestra.Col = 6
                            Muestra.Text = "SV"
                        Case 6
                            Muestra.Col = 6
                            Muestra.Text = "SVI"
                        Case 7
                            Muestra.Col = 6
                            Muestra.Text = "SVII"
                        Case Else
                    End Select
                    
                    Select Case Val(rstSolGuiaTotal!Hasta)
                        Case 1
                            Muestra.Col = 7
                            Muestra.Text = "SI"
                        Case 2
                            Muestra.Col = 7
                            Muestra.Text = "SII"
                        Case 3
                            Muestra.Col = 7
                            Muestra.Text = "SIII"
                        Case 4
                            Muestra.Col = 7
                            Muestra.Text = "SIV"
                        Case 5
                            Muestra.Col = 7
                            Muestra.Text = "SV"
                        Case 6
                            Muestra.Col = 7
                            Muestra.Text = "SVI"
                        Case 7
                            Muestra.Col = 7
                            Muestra.Text = "SVII"
                        Case Else
                    End Select
                    
                    Muestra.Col = 8
                    Muestra.Text = rstSolGuiaTotal!Observaciones
                    
                    Auxiliar(Renglon, 1) = rstSolGuiaTotal!Clave
                    Auxiliar(Renglon, 3) = rstSolGuiaTotal!Fecha
                    Auxiliar(Renglon, 4) = rstSolGuiaTotal!Tipo
                    Auxiliar(Renglon, 5) = rstSolGuiaTotal!Articulo
                    Auxiliar(Renglon, 6) = rstSolGuiaTotal!Terminado
                    Auxiliar(Renglon, 7) = rstSolGuiaTotal!Cantidad
                    Auxiliar(Renglon, 8) = rstSolGuiaTotal!Hasta
                    Auxiliar(Renglon, 9) = rstSolGuiaTotal!Observaciones
                    Auxiliar(Renglon, 10) = rstSolGuiaTotal!Desde
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        rstSolGuiaTotal.Close
    End If
    
    TotalSolicitud = Renglon
    
    For Cicla = 1 To TotalSolicitud
        WClave = Muestra.TextMatrix(Cicla, 3)
        spArticulo = "ConsultaArticulo " + "'" + WClave + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Muestra.Row = Cicla
            Muestra.Col = 4
            Muestra.Text = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        spTerminado = "ConsultaTerminado " + "'" + WClave + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            Muestra.Row = Cicla
            Muestra.Col = 4
            Muestra.Text = rstTerminado!Descripcion
            rstTerminado.Close
        End If
    Next Cicla
    
    Call Conecta_Empresa
    
    Renglon = Renglon + 1
    Muestra.Row = Renglon
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1

End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Solicitud"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Producto"
    
    Muestra.Col = 4
    Muestra.Text = "Descripcion"
    
    Muestra.Col = 5
    Muestra.Text = "Cantidad"
    
    Muestra.Col = 6
    Muestra.Text = "Origen"
    
    Muestra.Col = 7
    Muestra.Text = "Destino"
    
    Muestra.Col = 8
    Muestra.Text = "Observaciones"
    
    Muestra.Col = 9
    Muestra.Text = "Resp."
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub
