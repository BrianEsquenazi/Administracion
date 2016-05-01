VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgOrdPenDyII 
   Caption         =   "Listado de Ordenes Pendientes de Materia Prima"
   ClientHeight    =   6525
   ClientLeft      =   2640
   ClientTop       =   1020
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   6630
   Begin VB.CommandButton AvisoError 
      Caption         =   "No se puede emitir el reporte. Sistema sin Conexion con las otras plantas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      Picture         =   "ordpendyii.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      ItemData        =   "ordpendyii.frx":0742
      Left            =   120
      List            =   "ordpendyii.frx":0749
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2640
         TabIndex        =   6
         Top             =   960
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2640
         TabIndex        =   0
         Top             =   480
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
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WOrdPenDyII.rpt"
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
Attribute VB_Name = "PrgOrdPenDyII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstListaOrdenII As Recordset
Dim spListaOrdenII As String
Dim XParam As String
Dim Vector(10000, 2) As String
Dim Empe(100, 10) As String
Dim ListaOrden(10000, 15) As String
Dim ListaObserva(100, 2) As String

Private Sub Acepta_Click()

    WSalidaError = ""
    On Error GoTo Control_error
    
    XEmpresa = WEmpresa

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WEmpresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    On Error GoTo 0

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

    If WSalidaError = "N" Then Exit Sub


    Sql1 = "DELETE ListaOrdenII"
    spListaOrdenII = Sql1
    Set rstListaOrdenII = db.OpenRecordset(spListaOrdenII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    XEmpresa = WEmpresa
    
    
    LugarLista = 0
    Erase ListaOrden
    
    
    
    For CicloEmpresa = 1 To 2
    
        If CicloEmpresa = 1 Then
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        Erase Vector
        Lugar = 0

        XParam = "'" + "'"

        spOrden = "ModificaOrdenSaldo " + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
        Erase Vector
        Lugar = 0

        spOrden = "ListaOrdenTotal "
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                        WClave = rstOrden!Clave
                        WOrden = rstOrden!Orden
                        WFecha2 = rstOrden!fecha2
                        WSaldo = Str$(rstOrden!Cantidad - rstOrden!Recibida)
                        If Val(WSaldo) > 0 Then
                            Entra = "S"
                            For XX = 1 To Lugar
                                If Val(Vector(XX, 1)) = WOrden Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next XX
                            
                            If Entra = "S" Then
                                Lugar = Lugar + 1
                                Vector(Lugar, 1) = WOrden
                                Vector(Lugar, 2) = Right$(WFecha2, 4) + Mid$(WFecha2, 4, 2) + Left$(WFecha2, 2)
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
        
        For XX = 1 To Lugar
            WOrden = Vector(XX, 1)
            WFecha2 = Vector(XX, 2)
            XParam = "'" + WOrden + "','" _
                     + WFecha2 + "'"
    
            spOrden = "ModificaOrdenFecha2 " + XParam
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        Next XX
        
        Sql1 = "Select *"
        Sql2 = " FROM Orden"
        Sql3 = " Where Orden.Saldo > 0"
        Sql4 = " and Orden.Articulo >= " + "'" + Desde.Text + "'"
        Sql5 = " and Orden.Articulo <= " + "'" + Hasta.Text + "'"
        Sql6 = " Order by Orden.Clave"
        spOrden = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
                .MoveFirst
                Do
            
                    LugarLista = LugarLista + 1
                    
                    ZNumeroPedido = IIf(IsNull(rstOrden!PedidoImpo), "", rstOrden!PedidoImpo)
                    ZFechaPedido = IIf(IsNull(rstOrden!FechaImpo), "", rstOrden!FechaImpo)
                    ZTipoPedido = IIf(IsNull(rstOrden!TipoImpo), "", rstOrden!TipoImpo)
                    
                    ListaOrden(LugarLista, 1) = Str$(rstOrden!Orden)
                    ListaOrden(LugarLista, 2) = rstOrden!Fecha
                    ListaOrden(LugarLista, 3) = rstOrden!Proveedor
                    ListaOrden(LugarLista, 4) = rstOrden!Articulo
                    ListaOrden(LugarLista, 5) = Str$(rstOrden!Cantidad)
                    ListaOrden(LugarLista, 6) = rstOrden!fecha2
                    ListaOrden(LugarLista, 7) = Str$(rstOrden!Saldo)
                    ListaOrden(LugarLista, 8) = Str$(rstOrden!Carpeta)
                    ListaOrden(LugarLista, 9) = Str$(rstOrden!Renglon)
                    ListaOrden(LugarLista, 10) = ZNumeroPedido
                    ListaOrden(LugarLista, 11) = ZFechaPedido
                    ListaOrden(LugarLista, 12) = ZTipoPedido
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstOrden.Close
        
        End If
        
    Next CicloEmpresa
    
    
    
    
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
    
    
    
    For Ciclo = 1 To LugarLista
    
        WOrden = ListaOrden(Ciclo, 1)
        WFecha = ListaOrden(Ciclo, 2)
        WProveedor = ListaOrden(Ciclo, 3)
        WArticulo = ListaOrden(Ciclo, 4)
        WCantidad = ListaOrden(Ciclo, 5)
        WFechaEntrega = ListaOrden(Ciclo, 6)
        WSaldo = ListaOrden(Ciclo, 7)
        WCarpeta = ListaOrden(Ciclo, 8)
        WRenglon = ListaOrden(Ciclo, 9)
        WNumeroPedido = ListaOrden(Ciclo, 10)
        WFechaPedido = ListaOrden(Ciclo, 11)
        WTipoPedido = ListaOrden(Ciclo, 12)
        
        LugarObserva = 0
        Erase ListaObserva
        
        Sql1 = "Select *"
        Sql2 = " FROM ObservaOrden"
        Sql3 = " Where ObservaOrden.Carpeta = " + "'" + WCarpeta + "'"
        Sql4 = " Order by ObservaOrden.Clave"
        spObservaOrden = Sql1 + Sql2 + Sql3 + Sql4
        Set rstObservaOrden = db.OpenRecordset(spObservaOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstObservaOrden.RecordCount > 0 Then
            With rstObservaOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        LugarObserva = LugarObserva + 1
                        
                        ListaObserva(LugarObserva, 1) = rstObservaOrden!texto1
                        ListaObserva(LugarObserva, 2) = rstObservaOrden!texto2
                    
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstObservaOrden.Close
        End If
        
        If LugarObserva = 0 Then
            LugarObserva = 1
        End If
        
        For Cicla = 1 To LugarObserva
        
            WTexto1 = ListaObserva(Cicla, 1)
            WTexto2 = ListaObserva(Cicla, 2)
            WRenglonII = Str$(Cicla)
        
            Sql1 = "INSERT INTO ListaOrdenII ("
            Sql2 = "Clave ,"
            Sql3 = "Orden ,"
            Sql4 = "Renglon ,"
            Sql5 = "RenglonII ,"
            Sql6 = "Fecha ,"
            Sql7 = "Proveedor ,"
            Sql8 = "Articulo ,"
            Sql9 = "Cantidad ,"
            Sql10 = "FechaEntrega ,"
            Sql11 = "Saldo ,"
            Sql12 = "Carpeta ,"
            Sql13 = "Texto1 ,"
            Sql14 = "Texto2 ,"
            Sql15 = "PedidoImpo ,"
            Sql16 = "FechaImpo ,"
            Sql17 = "TipoImpo )"
            Sql18 = "Values ("
            Sql19 = "'" + WClave + "',"
            Sql20 = "'" + WOrden + "',"
            Sql21 = "'" + WRenglon + "',"
            Sql22 = "'" + WRenglonII + "',"
            Sql23 = "'" + WFecha + "',"
            Sql24 = "'" + WProveedor + "',"
            Sql25 = "'" + WArticulo + "',"
            Sql26 = "'" + WCantidad + "',"
            Sql27 = "'" + WFechaEntrega + "',"
            Sql28 = "'" + WSaldo + "',"
            Sql29 = "'" + WCarpeta + "',"
            Sql30 = "'" + WTexto1 + "',"
            Sql31 = "'" + Left$(WTexto2, 50) + "',"
            Sql32 = "'" + WNumeroPedido + "',"
            Sql33 = "'" + WFechaPedido + "',"
            Sql34 = "'" + WTipoPedido + "')"
        
            spListaOrdenII = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                       Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                       Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                       Sql31 + Sql32 + Sql33 + Sql34
            Set rstListaOrdenII = db.OpenRecordset(spListaOrdenII, dbOpenSnapshot, dbSQLPassThrough)
        
        Next Cicla
        
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Ordenes Pendientes por Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Orden.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaOrdenII.Orden, ListaOrdenII.RenglonII, ListaOrdenII.Fecha, ListaOrdenII.Articulo, ListaOrdenII.Cantidad, ListaOrdenII.FechaEntrega, ListaOrdenII.Saldo, ListaOrdenII.Carpeta, ListaOrdenII.Texto1, ListaOrdenII.Texto2, " _
                    + "Articulo.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.ListaOrdenII ListaOrdenII, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "ListaOrdenII.Articulo = Articulo.Codigo AND " _
                    + "ListaOrdenII.Proveedor = Proveedor.Proveedor "
    
    Rem Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
    Exit Sub
    
Control_error:
    Rem MsgBox Err.Description
    Beep
    AvisoError.Visible = True
    WSalidaError = "N"
    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgOrdPenDyII.Hide
    Unload Me
    Menu.Show
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
                        WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                        If WReventa = 1 Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                        End If
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
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstArticulo
            
            Indice = Pantalla.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                    Desde.Text = rstArticulo!Codigo
                    Hasta.Text = rstArticulo!Codigo
                End If
            End With
            Desde.SetFocus
        Case Else
    End Select
    
End Sub




Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spArticulo = "ListaArticuloConsulta"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
            
                WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                If WReventa = 1 Then
                    Da = Len(rstArticulo!Descripcion) - WEspacios
                    For Aaa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            Exit For
                        End If
                    Next Aaa
                End If
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



Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Sub Form_Load()

    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgOrdPenDyII.Caption = "Listado de Ordenes Pendientes de Materia Prima : " + !Nombre
        End If
    End With
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
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
        Desde.SetFocus
    End If
End Sub


