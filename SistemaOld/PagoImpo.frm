VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPagoImpo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyeccion de Pagos de Importaciones"
   ClientHeight    =   4785
   ClientLeft      =   2925
   ClientTop       =   2415
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2175
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WMovgascon.rpt"
      Destination     =   1
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
      Left            =   4680
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "PagoImpo.frx":0000
      Left            =   0
      List            =   "PagoImpo.frx":0007
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPagoImpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim XParam As String
Dim WPagos(1000, 2) As String
Dim Vector(1000, 20) As String
Dim WCalcula(1000, 10) As String
Dim WCarpe(10, 10) As String
Dim WEmpre(10) As String
Dim TotalEmpre As Integer
Dim EmpresaAnterior As String
Dim OtraEmpresa As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Proyeccion de Importaciones"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Posdat = "Desde el " + Desdefecha.Text + " hasta el " + HastaFecha.Text
            .Update
        End If
    End With
    
    EmpresaAnterior = WEmpresa
    
    For XX = 1 To 9
    
        Select Case Val(WEmpre(XX))
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
    
        spMovgas = "ModificaMovgasProceso0"
        Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
        
    Next XX
    
    For XX = 1 To TotalEmpre
    
        Select Case Val(WEmpre(XX))
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
    
        Erase Vector
        Entra = 0
    
        spOrden = "ListaOrdenTotal"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If rstOrden!Renglon = 1 Then
                            XTipo = IIf(IsNull(rstOrden!Tipo), "", rstOrden!Tipo)
                            If Val(XTipo) = 1 Then
                                WAno = Right$(rstOrden!Fecha2, 4)
                                WMes = Mid$(rstOrden!Fecha2, 4, 2)
                                WDia = Left$(rstOrden!Fecha2, 2)
                                WCompara = WAno + WMes + WDia
                                If WCompara >= WDesde And WCompara <= WHasta Then
                                    aa = rstOrden!Clave
                                    Entra = Entra + 1
                                    Vector(Entra, 1) = Str$(rstOrden!Orden)
                                    Vector(Entra, 2) = Str$(rstOrden!Carpeta)
                                    Vector(Entra, 3) = rstOrden!Fecha2
                                    Rem Vector(Entra, 4) = rstOrden!Empresa
                                End If
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
    
        For Cicla = 1 To Entra
    
            XOrden = Vector(Cicla, 1)
            XCarpeta = Vector(Cicla, 2)
            XFechaentrega = Vector(Cicla, 3)
            
            Rem calcula el costo flete
            
            XCostoFlete = 0
            Erase WCalcula
            Ingre = 0
        
            spOrden = "ListaOrden " + "'" + XOrden + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                With rstOrden
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Ingre = Ingre + 1
                            WCalcula(Ingre, 1) = rstOrden!Articulo
                            WCalcula(Ingre, 2) = Str$(rstOrden!Cantidad)
                            WCalcula(Ingre, 4) = Str$(rstOrden!Precio)
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstOrden.Close
            End If
    
            Rem For Ciclo = 1 To Ingre
            Rem     spArticulo = "ConsultaArticulo " + "'" + WCalcula(Ciclo, 1) + "'"
            Rem     Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            Rem     If rstArticulo.RecordCount > 0 Then
            Rem         WCalcula(Ciclo, 4) = Str$(rstArticulo!Flete)
            Rem         rstArticulo.Close
            Rem     End If
            Rem Next Ciclo
        
            For Ciclo = 1 To Ingre
                XCostoFlete = XCostoFlete + (Val(WCalcula(Ciclo, 4)) * Val(WCalcula(Ciclo, 2)))
            Next Ciclo
            
            Rem calcula los gastos
            
            XGastos = 0
            
            OtraEmpresa = WEmpresa
            Select Case Val(EmpresaAnterior)
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 8
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
            spMovgas = "Listamovgas " + "'" + XCarpeta + "'"
            Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovgas.RecordCount > 0 Then
                With rstMovgas
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            XGastos = XGastos + rstMovgas!Importe
                    
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstMovgas.Close
            End If
            
            XPagado = 0
            Erase WPagos
            Pag = 0
            
            spPagos = "ListaPagosCarpetaTotal "
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
                With rstPagos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Erase WCarpe
                                
                            WCarpe(1, 1) = IIf(IsNull(rstPagos!Carpeta), "0", rstPagos!Carpeta)
                            WCarpe(2, 1) = IIf(IsNull(rstPagos!Carpeta1), "0", rstPagos!Carpeta1)
                            WCarpe(3, 1) = IIf(IsNull(rstPagos!Carpeta2), "0", rstPagos!Carpeta2)
                            WCarpe(4, 1) = IIf(IsNull(rstPagos!Carpeta3), "0", rstPagos!Carpeta3)
                            WCarpe(5, 1) = IIf(IsNull(rstPagos!Carpeta4), "0", rstPagos!Carpeta4)
            
                            WCarpe(1, 2) = IIf(IsNull(rstPagos!ImpoCarpeta), "0", rstPagos!ImpoCarpeta)
                            WCarpe(2, 2) = IIf(IsNull(rstPagos!ImpoCarpeta1), "0", rstPagos!ImpoCarpeta1)
                            WCarpe(3, 2) = IIf(IsNull(rstPagos!ImpoCarpeta2), "0", rstPagos!ImpoCarpeta2)
                            WCarpe(4, 2) = IIf(IsNull(rstPagos!ImpoCarpeta3), "0", rstPagos!ImpoCarpeta3)
                            WCarpe(5, 2) = IIf(IsNull(rstPagos!ImpoCarpeta4), "0", rstPagos!ImpoCarpeta4)
                                
                            SumaCarpeta = 0
                            For XXX = 1 To 5
                                SumaCarpeta = SumaCarpeta + Val(WCarpe(XXX, 2))
                            Next XXX
                                
                            If SumaCarpeta = 0 And WCarpe(1, 1) <> 0 Then
                                WCarpe(1, 2) = rstPagos!Importe
                            End If
                                
                            For XXX = 1 To 5
                                If Val(XCarpeta) = Val(WCarpe(XXX, 1)) And Val(XCarpeta) <> 0 Then
                                    Pag = Pag + 1
                                    WPagos(Pag, 1) = rstPagos!Fecha
                                    WPagos(Pag, 2) = WCarpe(XXX, 2)
                                End If
                            Next XXX
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPagos.Close
            End If
            
            For XCiclo = 1 To Pag
            
                spCambios = "ConsultaCambio  " + "'" + WPagos(XCiclo, 1) + "'"
                Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                If rstCambios.RecordCount > 0 Then
                    Paridad = rstCambios!Cambio
                    rstCambios.Close
                        Else
                    Paridad = 1
                End If
                
                XPagado = XPagado + (Val(WPagos(XCiclo, 2)) / Paridad)
                
            Next XCiclo
            
            Rem actualiza archivo
        
            WCarpeta = XCarpeta
            WFechaLlegada = XFechaentrega
            WAno = Right$(XFechaentrega, 4)
            WMes = Mid$(XFechaentrega, 4, 2)
            WDia = Left$(XFechaentrega, 2)
            WCompara = WAno + WMes + WDia
            WOrdFechaLLegada = WCompara
            WCostoFlete = Str$(XCostoFlete)
            WGastos = Str$(XGastos)
            WPagado = Str$(XPagado)
            
            XParam = "'" + WCarpeta + "','" _
                        + WFechaLlegada + "','" _
                        + WOrdFechaLLegada + "','" _
                        + WCostoFlete + "','" _
                        + WGastos + "','" _
                        + WPagado + "'"
                         
            spMovgas = "ModificaMovgasProceso " + XParam
            Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
            
            Select Case Val(OtraEmpresa)
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
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
            
        Next Cicla
        
    Next XX
    
    Erase Vector
    Entra = 0
    
    Select Case Val(EmpresaAnterior)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    
    spMovgas = "ListaMovgasTotal"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovgas.RecordCount > 0 Then
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstMovgas!Costoflete <> 0 Or rstMovgas!gastos <> 0 Then
                    If rstMovgas!Renglon = 1 Then
                        Entra = Entra + 1
                        Vector(Entra, 1) = rstMovgas!Empresa
                        Vector(Entra, 2) = rstMovgas!Carpeta
                        Vector(Entra, 3) = rstMovgas!Fecha
                        Vector(Entra, 4) = Str$(rstMovgas!Derechos)
                        Vector(Entra, 5) = rstMovgas!Orden
                        Vector(Entra, 6) = rstMovgas!Concepto
                        Vector(Entra, 7) = Str$(rstMovgas!Importe)
                        Vector(Entra, 8) = rstMovgas!Auxiliar
                        Vector(Entra, 9) = rstMovgas!OrdFecha
                        Vector(Entra, 10) = rstMovgas!Proveedor
                        Vector(Entra, 11) = rstMovgas!Origen
                        Vector(Entra, 12) = rstMovgas!Moneda
                        Vector(Entra, 13) = ""
                        Vector(Entra, 14) = ""
                        Vector(Entra, 15) = rstMovgas!FechaLLegada
                        Vector(Entra, 16) = rstMovgas!OrdFechaLLegada
                        Vector(Entra, 17) = Str$(rstMovgas!Costoflete)
                        Vector(Entra, 18) = Str$(rstMovgas!gastos)
                        Vector(Entra, 19) = Str$(rstMovgas!Pagado)
                    End If
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovgas.Close
    End If
    
    spMovgascon = "BorrarMovGasCon "
    Set rstMovgascon = db.OpenRecordset(spMovgascon, dbOpenSnapshot, dbSQLPassThrough)
    
    For XX = 1 To Entra
    
        XEmpresa = Vector(XX, 1)
        XCarpeta = Vector(XX, 2)
        XFecha = Vector(XX, 3)
        Rem XDerechos = Vector(Xx, 4)
        XDerechos = ""
        XOrden = Vector(XX, 5)
        XConcepto = Vector(XX, 6)
        XImporte = Vector(XX, 7)
        XAuxiliar = Vector(XX, 8)
        XOrdFecha = Vector(XX, 9)
        XProveedor = Vector(XX, 10)
        XOrigen = Vector(XX, 11)
        XMoneda = Vector(XX, 12)
        XMarca = Vector(XX, 13)
        XImpoDerechos = Vector(XX, 14)
        XFechaLLegada = Vector(XX, 15)
        XOrdFechaLLegada = Vector(XX, 16)
        XCostoFlete = Vector(XX, 17)
        XGastos = Vector(XX, 18)
        XPagado = Vector(XX, 19)
        XClave = XEmpresa + XCarpeta
        
        XParam = "'" + XClave + "','" _
                     + XEmpresa + "','" + XCarpeta + "','" _
                     + XFecha + "','" + XDerechos + "','" _
                     + XOrden + "','" + XConcepto + "','" _
                     + XImporte + "','" + XAuxiliar + "','" _
                     + XOrdFecha + "','" + XProveedor + "','" _
                     + XOrigen + "','" + XMoneda + "','" _
                     + XMarca + "','" _
                     + XImpoDerecho + "','" _
                     + XFechaLLegada + "','" _
                     + XOrdFechaLLegada + "','" _
                     + XCostoFlete + "','" _
                     + XGastos + "','" _
                     + XPagado + "'"
                         
        spMovgascon = "AltaMovgasCon " + XParam
        Set rstMovgascon = db.OpenRecordset(spMovgascon, dbOpenSnapshot, dbSQLPassThrough)
        
    Next XX
    
    Listado.GroupSelectionFormula = "{MovGasCon.Carpeta} in 0 to 999999 and {MovGasCon.Empresa} in 0 to 9999 and {@Suma} > 0.00"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT MovGasCon.Empresa, MovGasCon.Carpeta, MovGasCon.Fecha, MovGasCon.Orden, MovGasCon.Origen, MovGasCon.Moneda, MovGasCon.FechaLLegada, MovGasCon.OrdFechaLLegada, MovGasCon.CostoFlete, MovGasCon.Gastos, MovGasCon.Pagado " _
                        + "From " _
                        + DSQ + ".dbo.MovGasCon MovGasCon " _
                        + "Where " _
                        + "MovGasCon.Empresa >= 0 AND MovGasCon.Empresa <= 9999 AND " _
                        + "MovGasCon.Carpeta >= 0 AND MovGasCon.Carpeta <= 999999"
    
    Listado.DataFiles(1) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    Desdefecha.SetFocus
    PrgPagoImpo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Erase WEmpre
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7
            WEmpre(1) = "0001"
            WEmpre(2) = "0003"
            WEmpre(3) = "0005"
            WEmpre(4) = "0006"
            WEmpre(5) = "0007"
            TotalEmpre = 5
        Case Else
            WEmpre(1) = "0002"
            WEmpre(2) = "0004"
            WEmpre(3) = "0008"
            WEmpre(4) = "0009"
            TotalEmpre = 4
    End Select
    
End Sub

