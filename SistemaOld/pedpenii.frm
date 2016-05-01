VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgPedPenII 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pedidos Pendientes"
   ClientHeight    =   4605
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3975
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Mensaje2 
         Alignment       =   2  'Center
         Caption         =   "Productos a Emitir Notas de Credito"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   3
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Mensaje1 
         Alignment       =   2  'Center
         Caption         =   "Pedidos a Retirar por el cliente"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
Attribute VB_Name = "PrgPedPenII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Tres As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstLiberaTerminado As Recordset
Dim spLiberaTerminado As String
Dim XParam As String
Dim WVector(1000, 5) As String
Dim LugarVector As Integer
Dim WTipopro As String
Dim WDesdeFec As String
Dim WHastaFec As String
Dim LugarLibera As Integer
Dim LugarLiberaI As Integer
Dim LugarLiberaII As Integer
Dim LugarLiberaIII As Integer

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)


    WDesdeFec = "01/01/2005"
    WHastaFec = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)

    WAno = Right$(WDesdeFec, 4)
    WMes = Mid$(WDesdeFec, 4, 2)
    WDia = Left$(WDesdeFec, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(WHastaFec, 4)
    WMes = Mid$(WHastaFec, 4, 2)
    WDia = Left$(WHastaFec, 2)
    WHasta = WAno + WMes + WDia

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Sql1 = "UPDATE Pedido SET "
    Sql2 = " Importe = Cantidad - Facturado "
    Sql3 = " Where OrdFecEntrega >= " + "'" + WDesde + "'"
    Sql4 = " and OrdFecEntrega <= " + "'" + WHasta + "'"
    Sql5 = " and TipoPed = 4"

    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Erase WVector
    LugarVector = 0
    
    spPedido = "ListaPedidoPend"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstPedido!Autorizo <> "N" Then
                        EntraVector = "S"
                        For Ciclo = 1 To LugarVector
                            If WVector(Ciclo, 1) = rstPedido!Terminado Then
                                EntraVector = "N"
                                Exit For
                            End If
                        Next Ciclo
                        If EntraVector = "S" Then
                            LugarVector = LugarVector + 1
                            WVector(LugarVector, 1) = rstPedido!Terminado
                        End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To LugarVector
        WProducto = WVector(Ciclo, 1)
        WTipopro = Left$(WProducto, 2)
        Select Case WTipopro
            Case "DY", "DW"
                WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                XParam = "'" + WArticulo + "','" _
                        + WDescripcion + "'"
                spPedido = "ModificaPedidoArticulo " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
                WTerminado = WProducto
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                
                XParam = "'" + WTerminado + "','" _
                        + WDescripcion + "'"
                spPedido = "ModificaPedidoTerminado " + XParam
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        End Select
    Next Ciclo
    
    LugarLibera = 0
    LugarLiberaI = 0
    LugarLiberaII = 0
    LugarLiberaIII = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM LiberaTerminado"
    Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
    Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
    spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstLiberaTerminado.RecordCount > 0 Then
        LugarLibera = 1
        LugarLiberaI = 1
        rstLiberaTerminado.Close
    End If

    WEmpresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Sql1 = "Select *"
    Sql2 = " FROM LiberaTerminado"
    Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
    Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
    spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstLiberaTerminado.RecordCount > 0 Then
        LugarLibera = 1
        LugarLiberaII = 1
        rstLiberaTerminado.Close
    End If

    WEmpresa = "0007"
    txtOdbc = "Empresa07"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Sql1 = "Select *"
    Sql2 = " FROM LiberaTerminado"
    Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
    Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
    spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
    Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstLiberaTerminado.RecordCount > 0 Then
        LugarLibera = 1
        LugarLiberaIII = 1
        rstLiberaTerminado.Close
    End If

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WEntra = "N"
    
    If LugarVector > 0 Then
        Mensaje1.Visible = True
        WEntra = "S"
            Else
        Mensaje1.Visible = False
    End If
        
    If LugarLibera > 0 Then
        Mensaje2.Visible = True
        WEntra = "S"
            Else
        Mensaje2.Visible = False
    End If
    
    If WEntra = "S" Then
        PrgPedPenII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgPedPenII.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    If LugarVector > 0 Then

        Listado.WindowTitle = "Listado de Pedidos Pendientes"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
   
        Listado.Destination = 1
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.FechaOrd, Pedido.Facturado, Pedido.Importe, Pedido.Autorizo, Pedido.Tipoped, Pedido.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.Importe > 0 AND " _
                    + "Pedido.Autorizo <> 'N'"
    
        Listado.DataFiles(2) = WEmpresa + "auxi.mdb"
        Listado.Connect = Connect()

        Listado.Action = 1
    
    End If
    
    If LugarLiberaI > 0 Then

        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Listado.ReportFileName = ""

        Listado.WindowTitle = "Listado de Pedidos Pendientes"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
   
        Listado.ReportFileName = "LISTAFACTURAR.RPT"
        Listado.Destination = 1
         Rem BY NAN
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT LiberaTerminado.Producto, LiberaTerminado.Fecha, LiberaTerminado.Partida, LiberaTerminado.Marca, LiberaTerminado.Cliente, LiberaTerminado.Cantidad, LiberaTerminado.Observa, LiberaTerminado.Tipo, LiberaTerminado.ImpreVentas, LiberaTerminado.Pedidodevol " _
                    + "From " _
                    + DSQ + ".dbo.LiberaTerminado LiberaTerminado " _
                    + "Where " _
                    + "LiberaTerminado.Cliente > '' AND " _
                    + "LiberaTerminado.ImpreVentas = 'N'"
        Listado.Connect = Connect()

        Listado.Action = 1
        
        Rem Sql1 = "UPDATE LiberaTerminado SET "
        Rem Sql2 = " ImpreVentas = " + "'" + "S" + "'"
        Rem Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
        Rem Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
        
        Rem spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
        Rem Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)

        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If
        
    If LugarLiberaII > 0 Then

        WEmpresa = "0005"
        txtOdbc = "Empresa05"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Listado.ReportFileName = ""

        Listado.WindowTitle = "Listado de Pedidos Pendientes"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
   
        Listado.ReportFileName = "LISTAFACTURAR.RPT"
        Listado.Destination = 1
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT LiberaTerminado.Producto, LiberaTerminado.Fecha, LiberaTerminado.Partida, LiberaTerminado.Marca, LiberaTerminado.Cliente, LiberaTerminado.Cantidad, LiberaTerminado.Observa, LiberaTerminado.Tipo, LiberaTerminado.ImpreVentas, LiberaTerminado.PedidoDevol " _
                    + "From " _
                    + DSQ + ".dbo.LiberaTerminado LiberaTerminado " _
                    + "Where " _
                    + "LiberaTerminado.Cliente > '' AND " _
                    + "LiberaTerminado.ImpreVentas = 'N'"
        Listado.Connect = Connect()

        Listado.Action = 1
        
        Rem Sql1 = "UPDATE LiberaTerminado SET "
        Rem Sql2 = " ImpreVentas = " + "'" + "S" + "'"
        Rem Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
        Rem Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
        
        Rem spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
        Rem Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)

        Rem WEmpresa = "0001"
        Rem txtOdbc = "Empresa01"
        Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If
        
    If LugarLiberaIII > 0 Then

        WEmpresa = "0007"
        txtOdbc = "Empresa07"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        Listado.ReportFileName = ""

        Listado.WindowTitle = "Listado de Pedidos Pendientes"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
   
        Listado.ReportFileName = "LISTAFACTURAR.RPT"
        Listado.Destination = 1
        Rem Listado.Destination = 0
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT LiberaTerminado.Producto, LiberaTerminado.Fecha, LiberaTerminado.Partida, LiberaTerminado.Marca, LiberaTerminado.Cliente, LiberaTerminado.Cantidad, LiberaTerminado.Observa, LiberaTerminado.Tipo, LiberaTerminado.ImpreVentas, LiberaTerminado.PedidoDevol " _
                    + "From " _
                    + DSQ + ".dbo.LiberaTerminado LiberaTerminado " _
                    + "Where " _
                    + "LiberaTerminado.Cliente > '' AND " _
                    + "LiberaTerminado.ImpreVentas = 'N'"
        Listado.Connect = Connect()

        Listado.Action = 1
        
        Rem Sql1 = "UPDATE LiberaTerminado SET "
        Rem Sql2 = " ImpreVentas = " + "'" + "S" + "'"
        Rem Sql3 = " Where LiberaTerminado.ImpreVentas = " + "'" + "N" + "'"
        Rem Sql4 = " and LiberaTerminado.Cliente <> " + "'" + "" + "'"
        
        Rem spLiberaTerminado = Sql1 + Sql2 + Sql3 + Sql4
        Rem Set rstLiberaTerminado = db.OpenRecordset(spLiberaTerminado, dbOpenSnapshot, dbSQLPassThrough)
        

        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    End If
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgPedPenII.Hide
    Unload Me
    Close
    End
End Sub


