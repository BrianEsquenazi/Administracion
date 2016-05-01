VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAvisoVtoSII 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Vencimiento de M.P."
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Impre 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "EXISTEN MATERIAS PRIMAS EN SI A VERIFICAR SU ESTADO"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
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
Attribute VB_Name = "PrgAvisoVtoSII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim ZArti(10000, 10) As String
Dim Empe(10, 10) As String
Dim WSaldo As Double
Dim ZFecha As String
Dim ZFechaVto As String
Dim XMes As String
Dim XAno As String

Dim ZDias As String
Dim ZComparaI As Date
Dim ZComparaII As Date

Dim ZLaudo As String
Dim ZOrdFecha As String
Dim ZArticulo As String
Dim ZCantidad As String
Dim ZSaldo As String
Dim ZVto As String
Dim ZDesEmpresa As String
Dim ZTitulo As String
Dim ZEmpresa As String
Dim XEmpresa As String

Private Sub Acepta_Click()

    OPEN_FILE_Empresa
    
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = "DELETE ListaVencimiento"
    spListaVencimiento = ZSql
    Set rstListaVencimiento = db.OpenRecordset(spListaVencimiento, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            ZDesEmpresa = !Nombre
        End If
    End With
    
    Erase ZArti
    ZLugar = 0
    ZSalida = "N"
    
    ZTitulo = "Planta SII   Fecha : " + Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Rem PROCESA LOS LAUDOS
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Where Saldo <> 0"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                        Else
                    
                    
                    
                    WArticulo = rstLaudo!Articulo
                    WCantidad = rstLaudo!Liberada
                    WFecha = rstLaudo!Fecha
                    WLaudo = rstLaudo!Laudo
                    WPartiOri = rstLaudo!partiori
                    WOrden = rstLaudo!Orden
                    WDevuelta = IIf(IsNull(rstLaudo!Devuelta), "0", rstLaudo!Devuelta)
                    WRechazo = IIf(IsNull(rstLaudo!Rechazo), "0", rstLaudo!Rechazo)
                    WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                    WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                    Call Redondeo(WSaldo)
                    WVencimiento = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                    WOrdVencimiento = IIf(IsNull(rstLaudo!OrdFechaVencimiento), "", rstLaudo!OrdFechaVencimiento)
                    
                    If WSaldo <> 0 Then
                    
                        Rem If WArticulo = "AA-133-100" Then Stop
                        
                        ZLugar = ZLugar + 1
                        ZArti(ZLugar, 1) = WLaudo
                        ZArti(ZLugar, 2) = WArticulo
                        ZArti(ZLugar, 3) = Str$(WCantidad)
                        ZArti(ZLugar, 4) = Str$(WSaldo)
                        Select Case Val(WEmpresa)
                            Case 1
                                ZArti(ZLugar, 5) = "Pta SI"
                            Case 2
                                ZArti(ZLugar, 5) = "Pta PI"
                            Case 3
                                ZArti(ZLugar, 5) = "Pta SII"
                            Case 4
                                ZArti(ZLugar, 5) = "Pta PII"
                            Case 5
                                ZArti(ZLugar, 5) = "Pta SIII"
                            Case 6
                                ZArti(ZLugar, 5) = "Pta SVI"
                            Case 7
                                ZArti(ZLugar, 5) = "Pta SV"
                            Case 8
                                ZArti(ZLugar, 5) = "Pta PIII"
                            Case 9
                                ZArti(ZLugar, 5) = "Pta PIV"
                            Case 10
                                ZArti(ZLugar, 5) = "Trabajo"
                            Case Else
                        End Select
                        ZArti(ZLugar, 6) = WVencimiento
                        ZArti(ZLugar, 7) = WOrdVencimiento
                        ZArti(ZLugar, 8) = WFecha
                    
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
    
    
    Rem PROCESA LAS GUIAS DE TRASLADO INTERNOS
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Guia"
    ZSql = ZSql + " Where Saldo <> 0"
    spMovguia = ZSql
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
                        
                    If rstMovguia!Tipo = "M" Then
                    
                        Rem If WArticulo = "AA-133-100" Then Stop
                        
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        WLaudo = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                        Call Redondeo(WSaldo)
                        
                        If WSaldo <> 0 Then
                        
                            ZLugar = ZLugar + 1
                            ZArti(ZLugar, 1) = WLaudo
                            ZArti(ZLugar, 2) = WArticulo
                            ZArti(ZLugar, 3) = Str$(WCantidad)
                            ZArti(ZLugar, 4) = Str$(WSaldo)
                            ZArti(ZLugar, 5) = ""
                            ZArti(ZLugar, 6) = ""
                            ZArti(ZLugar, 7) = ""
                            ZArti(ZLugar, 8) = ""
                    
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
    
    
    For Ciclo = 1 To ZLugar
    
        ZVto = ""
        
        ZLaudo = ZArti(Ciclo, 1)
        ZArticulo = ZArti(Ciclo, 2)
        ZCantidad = ZArti(Ciclo, 3)
        ZSaldo = ZArti(Ciclo, 4)
        ZEmpresa = ZArti(Ciclo, 5)
        ZFechaVto = ZArti(Ciclo, 6)
        ZFecha = ZArti(Ciclo, 8)
        
        If Trim(ZEmpresa) = "" Then
        
            XEmpresa = WEmpresa
    
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7
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
                
                Case Else
                    Empe(1, 1) = "0002"
                    Empe(1, 2) = "Empresa02"
                    Empe(2, 1) = "0004"
                    Empe(2, 2) = "Empresa04"
                    Empe(3, 1) = "0008"
                    Empe(3, 2) = "Empresa08"
                    Empe(4, 1) = "0009"
                    Empe(4, 2) = "Empresa09"
                    XHasta = 4
            End Select
    
            For Ciclo2 = 1 To XHasta
    
                WEmpresa = Empe(Ciclo2, 1)
                txtOdbc = Empe(Ciclo2, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo >= " + "'" + ZLaudo + "'"
                ZSql = ZSql + " and Articulo <= " + "'" + ZArticulo + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                        
                    ZFecha = rstLaudo!Fecha
                    ZFechaVto = IIf(IsNull(rstLaudo!FechaVencimiento), "", rstLaudo!FechaVencimiento)
                    
                    Select Case Val(WEmpresa)
                        Case 1
                            ZEmpresa = "Pta SI"
                        Case 2
                            ZEmpresa = "Pta PI"
                        Case 3
                            ZEmpresa = "Pta SII"
                        Case 4
                            ZEmpresa = "Pta PII"
                        Case 5
                            ZEmpresa = "Pta SIII"
                        Case 6
                            ZEmpresa = "Pta SVI"
                        Case 7
                            ZEmpresa = "Pta SV"
                        Case 8
                            ZEmpresa = "Pta PIII"
                        Case 9
                            ZEmpresa = "Pta PIV"
                        Case 10
                            ZEmpresa = "Trabajo"
                        Case Else
                    End Select
                        
                    rstLaudo.Close
                    Exit For
        
                End If
            
            Next Ciclo2
            
            Call Conecta_Empresa
        
        End If
                    
                    
        ZOrdFecha = Right$(ZFecha, 4) + Mid$(ZFecha, 4, 2) + Left$(ZFecha, 2)
        If ZFechaVto <> "" And ZFechaVto <> "  /  /    " Then
            Call Valida_fecha(ZFechaVto, Auxi)
            If Auxi = "S" Then
                ZVto = ZFechaVto
            End If
        End If
        
        If ZVto = "" Then
        
            ZMeses = 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZMeses = rstArticulo!meses
                rstArticulo.Close
            End If
        
            WMes = Val(Mid$(ZFecha, 4, 2))
            WAno = Val(Right$(ZFecha, 4))
            For ZCiclo = 1 To ZMeses
                WMes = WMes + 1
                If WMes > 12 Then
                    WAno = WAno + 1
                    WMes = 1
                End If
            Next ZCiclo
            
            XMes = Str$(WMes)
            XAno = Str$(WAno)
            Call Ceros(XMes, 2)
            Call Ceros(XAno, 4)
            If Val(Left$(ZFecha, 2)) <= 30 Then
                If Val(XMes) = 2 And Val(Left$(ZFecha, 2)) > 28 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = Left$(ZFecha, 3) + XMes + "/" + XAno
                End If
                    Else
                If Val(XMes) = 2 Then
                    ZVto = "28/" + XMes + "/" + XAno
                        Else
                    ZVto = "30/" + XMes + "/" + XAno
                End If
            End If
            
        End If
        
        ZComparaI = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZComparaII = ZVto
        
        ZDias = DateDiff("d", ZComparaI, ZComparaII)
        
        If Val(ZDias) < 30 Then
        
            ZSalida = "S"
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ListaVencimiento ("
            ZSql = ZSql + "Laudo ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Liberada ,"
            ZSql = ZSql + "Saldo ,"
            ZSql = ZSql + "FechaVencimiento ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "DesEmpresa ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Origen )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZLaudo + "',"
            ZSql = ZSql + "'" + ZFecha + "',"
            ZSql = ZSql + "'" + ZOrdFecha + "',"
            ZSql = ZSql + "'" + ZArticulo + "',"
            ZSql = ZSql + "'" + ZCantidad + "',"
            ZSql = ZSql + "'" + ZSaldo + "',"
            ZSql = ZSql + "'" + ZVto + "',"
            ZSql = ZSql + "'" + ZDias + "',"
            ZSql = ZSql + "'" + ZDesEmpresa + "',"
            ZSql = ZSql + "'" + ZTitulo + "',"
            ZSql = ZSql + "'" + ZEmpresa + "')"
        
            spListaVencimiento = ZSql
            Set rstListaVencimiento = db.OpenRecordset(spListaVencimiento, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
        
    If ZSalida = "S" Then
        PrgAvisoVtoSII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgAvisoVtoSII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Cancela_click()
    PrgAvisoVtoSII.Hide
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub


Private Sub Conecta_Empresa()

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
        Case Else
    End Select

End Sub

Private Sub Impre_Click()

    Listado.WindowTitle = "Verificacion de Vencimientos de Materia Prima"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaVencimiento.Laudo, ListaVencimiento.Fecha, ListaVencimiento.Articulo, ListaVencimiento.Liberada, ListaVencimiento.Saldo, ListaVencimiento.FechaVencimiento, ListaVencimiento.Dias, ListaVencimiento.DesEmpresa, ListaVencimiento.Titulo, ListaVencimiento.Origen, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.ListaVencimiento ListaVencimiento, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "ListaVencimiento.Articulo = Articulo.Codigo AND " _
            + "ListaVencimiento.Laudo >= 0 AND " _
            + "ListaVencimiento.Laudo <= 999999"
    Listado.Connect = Connect()
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "WListaVencimiento.rpt"
    
    Listado.Action = 1
    Call Cancela_click

End Sub
