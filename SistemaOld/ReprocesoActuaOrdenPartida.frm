VERSION 5.00
Begin VB.Form PrgReprocesoActuaOrdenPartida 
   AutoRedraw      =   -1  'True
   Caption         =   "Actualizacion de Costos de Importacion"
   ClientHeight    =   3750
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3750
   ScaleWidth      =   8145
   Begin VB.HScrollBar Barra 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox Contador 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Salida"
      Height          =   975
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Proceso"
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "PrgReprocesoActuaOrdenPartida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim XParamII As String
Dim Vector(100, 20) As String
Dim Gastos(100, 10) As String

Dim rstCarpeta As Recordset
Dim spCarpeta As String
Dim rstMovgas As Recordset
Dim spMovgas As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCostoPartida As Recordset
Dim spCostoPartida As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstCambios As Recordset
Dim spCambios As String

Dim WArancel As String
Dim WCarpeta As String
Dim WCosto As Double
Dim XCosto1 As Double
Dim XCosto2 As Double
Dim XCosto3 As Double
Dim ZCostoCompara As Double
Dim WLeyenda As Integer
Dim CargaEmpresa(12, 2) As String

Dim ZLugar As Integer
Dim ZVector(30000, 8) As String
Dim ZCarpeta As String

Dim WWEmpresa As String
Dim WWOrden As String
Dim WWArticulo As String
Dim WWCarpeta As String
Dim WWCosto As String
Dim WWLaudo As String

Dim ZZLaudo As String
Dim ZZOrden As String
Dim ZZTipoOrden As String
Dim ZZCarpeta As String
Dim ZZPrecio As String
Dim ZZMoneda As String
Dim ZZFecha As String
Dim ZZParidad As Double
Dim ZZArticulo As String

Private Sub Calcula_Importacion()

    WCarpeta = ZZCarpeta
    Call Ceros(WCarpeta, 6)
    
    WLaudo = ZZLaudo
    
    EmpresaAnterior = WEmpresa
    
    ZEmpresa = WEmpresa
    WOrden = ZZOrden
 
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Rem WEmpresa = "0008"
    Rem txtOdbc = "Empresa08"
    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
 
    Renglon = 0
    Erase Gastos
    ImpoGastos = 0
    ImpoSeguro = 0
    ImpoFlete = 0
    
    spMovgas = "ListaMovgas " + "'" + WCarpeta + "'"
    Set rstMovgas = db.OpenRecordset(spMovgas, dbOpenSnapshot, dbSQLPassThrough)

    If rstMovgas.RecordCount > 0 Then
    
        With rstMovgas
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstMovgas!concepto <> 10 Then
                        WArancel = Str$(rstMovgas!Derechos)
                        WOrden = Str$(rstMovgas!Orden)
                        WEmpresaOtro = rstMovgas!Empresa
                        Select Case rstMovgas!concepto
                            Case 2
                                ImpoSeguro = ImpoSeguro + rstMovgas!Importe
                            Case 4, 5
                                ImpoFlete = ImpoFlete + rstMovgas!Importe
                            Case Else
                                Renglon = Renglon + 1
                                Gastos(Renglon, 1) = Str$(rstMovgas!concepto)
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
    
    Select Case Val(ZEmpresa)
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
    
    ZCoeParidad = 1
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Orden = " + "'" + WOrden + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        ZNroInforme = rstInforme!Informe
        ZFechaInforme = rstInforme!Fecha
        rstInforme.Close
        spCambios = "ConsultaCambio  " + "'" + ZFechaInforme + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            ZParidad = rstCambios!Cambio
            ZParidadII = IIf(IsNull(rstCambios!CambioII), "0", rstCambios!CambioII)
            If ZParidadII <> 0 And ZParidad <> 0 Then
                ZCoeParidad = ZParidadII / ZParidad
            End If
            rstCambios.Close
        End If
    End If

    
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
                    
                    WMoneda = rstOrden!Moneda
                    ZZPrecio = rstOrden!Precio
                    If WMoneda = 2 Then
                        ZZPrecio = ZZPrecio * ZCoeParidad
                    End If
                    
                    Vector(Renglon, 1) = WCarpeta
                    Vector(Renglon, 2) = rstOrden!Articulo
                    Vector(Renglon, 3) = Str$(rstOrden!Cantidad)
                    Vector(Renglon, 4) = Str$(ZZPrecio)
                    Vector(Renglon, 5) = Str$(rstOrden!Cantidad * ZZPrecio)
                    Vector(Renglon, 6) = Str$(rstOrden!Derechos)
                    Vector(Renglon, 7) = ""
                    Vector(Renglon, 8) = ""
                    Vector(Renglon, 9) = ""
                    Vector(Renglon, 10) = ""
                    Vector(Renglon, 11) = WCarpeta + "01"
                    Vector(Renglon, 12) = Str$(ZZPrecio)
                    
                    WTotalImpo = WTotalImpo + (rstOrden!Cantidad * ZZPrecio)
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
    
    For Ciclo = 1 To Renglon
        WSeguro = 0
        WFlete = 0
        If ImpoSeguro <> 0 Then
            If WTotalImpo <> 0 And Val(Vector(Ciclo, 5)) <> 0 Then
                WSeguro = ((Val(Vector(Ciclo, 5)) / WTotalImpo) * ImpoSeguro) / Val(Vector(Ciclo, 3))
            End If
        End If
        If ImpoFlete <> 0 Then
            If WTotalPeso <> 0 And Val(Vector(Ciclo, 3)) <> 0 Then
                WFlete = ((Val(Vector(Ciclo, 3)) / WTotalPeso) * ImpoFlete) / Val(Vector(Ciclo, 3))
            End If
        End If
        WCosto = Val(Vector(Ciclo, 4)) + WSeguro + WFlete
        Rem WCosto = Val(Vector(Ciclo, 4))
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

    For Ciclo = 1 To Renglon
    
        WArticulo = Vector(Ciclo, 2)
        
        XCosto1 = Val(Vector(Ciclo, 9))
        XCosto2 = Val(Vector(Ciclo, 9)) * 1.03
        XCosto3 = Val(Vector(Ciclo, 9))
        
        CostoImpo = Val(Vector(Ciclo, 12))
        If CostoImpo <> 0 Then
            ZZCoeficiente = XCosto1 / CostoImpo
                Else
            ZZCoeficiente = 0
        End If
        
        Call Redondeo(XCosto1)
        Call Redondeo(XCosto2)
        Call Redondeo(XCosto3)
        
        WCosto1 = Str$(XCosto1)
        WCosto2 = Str$(XCosto2)
        WCosto3 = Str$(XCosto3)
        WFlete = Str$(CostoImpo)
        
        If WLeyenda > 0 Then
            XLeyenda = Str$(WLeyenda - 1)
                Else
            XLeyenda = "0"
        End If
        
        WWEmpresa = ZEmpresa
        WWOrden = WOrden
        WWArticulo = WArticulo
        WWCarpeta = WCarpeta
        WWCosto = WCosto1
        WWLaudo = WLaudo
        
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWLaudo, 6)
        
        WWClave = WWEmpresa + WWLaudo + WWArticulo
                
        If WWArticulo = ZZArticulo Then
                
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            Rem WEmpresa = "0008"
            Rem txtOdbc = "Empresa08"
            Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CostoPartida ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Carpeta ,"
            ZSql = ZSql + "Laudo ,"
            ZSql = ZSql + "Costo )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WWClave + "',"
            ZSql = ZSql + "'" + WWEmpresa + "',"
            ZSql = ZSql + "'" + WWOrden + "',"
            ZSql = ZSql + "'" + WWArticulo + "',"
            ZSql = ZSql + "'" + WWCarpeta + "',"
            ZSql = ZSql + "'" + WWLaudo + "',"
            ZSql = ZSql + "'" + WWCosto + "')"
            
            spCostoPartida = ZSql
            Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Next Ciclo
    
    Select Case Val(ZEmpresa)
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

End Sub



Private Sub Graba_Orden()

    WLaudo = ZZLaudo
    
    ZEmpresa = WEmpresa
    WOrden = ZZOrden
 
    Renglon = 0
    Erase Vector
    
    spOrden = "ListaOrden " + "'" + WOrden + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    
                    Vector(Renglon, 1) = rstOrden!Articulo
                    If ZZParidad <> 0 Then
                        Vector(Renglon, 2) = Str$(rstOrden!Precio / ZZParidad)
                            Else
                        Vector(Renglon, 2) = Str$(rstOrden!Precio)
                    End If
                    Vector(Renglon, 3) = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    WTotal = 0
    
    For Ciclo = 1 To Renglon
    
        WWEmpresa = ZEmpresa
        WWOrden = WOrden
        WWArticulo = Vector(Ciclo, 1)
        WWCarpeta = ""
        WWCosto = Vector(Ciclo, 2)
        WWLaudo = WLaudo
    
        Call Ceros(WWEmpresa, 2)
        Call Ceros(WWLaudo, 6)
        
        WWClave = WWEmpresa + WWLaudo + WWArticulo
        
        If WWArticulo = ZZArticulo Then
        
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            Rem WEmpresa = "0008"
            Rem txtOdbc = "Empresa08"
            Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CostoPartida"
            ZSql = ZSql + " Where CostoPartida.Clave = " + "'" + WWClave + "'"
            spCostoPartida = ZSql
            Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
            If rstCostoPartida.RecordCount > 0 Then
            
                rstCostoPartida.Close
                
                    Else
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO CostoPartida ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Empresa ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Carpeta ,"
                ZSql = ZSql + "Laudo ,"
                ZSql = ZSql + "Costo )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWEmpresa + "',"
                ZSql = ZSql + "'" + WWOrden + "',"
                ZSql = ZSql + "'" + WWArticulo + "',"
                ZSql = ZSql + "'" + WWCarpeta + "',"
                ZSql = ZSql + "'" + WWLaudo + "',"
                ZSql = ZSql + "'" + WWCosto + "')"
                
                spCostoPartida = ZSql
                Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Next Ciclo
    
    Select Case Val(ZEmpresa)
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

End Sub

Private Sub Cancela_click()
    PrgReprocesoActuaOrdenPartida.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Command1_Click()

    Stop

    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE CtaCte"
    Rem ZSql = ZSql + " Where OrdFecha < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Saldo = 0 "
    Rem spCtacte = ZSql
    Rem Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE CtaCtePrv"
    Rem ZSql = ZSql + " Where OrdFecha < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Saldo = 0 "
    Rem spCtactePrv = ZSql
    Rem Set RstCtaCtePrv = db.OpenRecordset(spCtactePrv, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Depositos"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spDeposito = ZSql
    Rem Set RstDeposito = db.OpenRecordset(spDeposito, dbOpenSnapshot, dbSQLPassThrough)
        
            
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Estadistica"
    Rem ZSql = ZSql + " Where OrdFecha < " + "'" + "20030101" + "'"
    Rem spEstadistica = ZSql
    Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Guia"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Saldo = 0 "
    Rem spGuia = ZSql
    Rem Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Hoja"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Saldo = 0 "
    Rem spHoja = ZSql
    Rem Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)

    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Imputac"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spImputac = ZSql
    Rem Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)

    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Informe"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spInforme = ZSql
    Rem Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Laudo"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Saldo = 0 "
    Rem spLaudo = ZSql
    Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Movlab"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spMovlab = ZSql
    Rem Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Movvar"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spMovvar = ZSql
    Rem Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Orden"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spOrden = ZSql
    Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Pagos"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spPago = ZSql
    Rem Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Pedido"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem spPedido = ZSql
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Recibos"
    Rem ZSql = ZSql + " Where FechaOrd < " + "'" + "20030101" + "'"
    Rem ZSql = ZSql + " and Estado2 <> 'P'"
    Rem spRecibo = ZSql
    Rem Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Proceso_Click()

    ZLugar = 0
    Erase ZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Order by Laudo.Clave"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstLaudo!Orden <> 0 And rstLaudo!Laudo <> 0 Then
                    
                    If rstLaudo!FechaOrd > "20031231" Then
                    
                        ZLugar = ZLugar + 1
                        
                        ZVector(ZLugar, 1) = Str$(rstLaudo!Laudo)
                        ZVector(ZLugar, 2) = Str$(rstLaudo!Orden)
                        ZVector(ZLugar, 3) = ""
                        ZVector(ZLugar, 4) = ""
                        ZVector(ZLugar, 5) = ""
                        ZVector(ZLugar, 6) = ""
                        ZVector(ZLugar, 7) = ""
                        ZVector(ZLugar, 8) = rstLaudo!Articulo
                        
                    End If
                    
                    End If
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZZLaudo = ZVector(Ciclo, 1)
        ZZOrden = ZVector(Ciclo, 2)
        ZZArticulo = ZVector(Ciclo, 8)
        
        Call Ceros(ZZOrden, 6)
        ZZClave = ZZOrden + "01"
    
        ZSql = ""
        ZSql = ZSql + "Select Orden.Orden, Orden.Tipo, Orden.Carpeta, Orden.Precio, Orden.Moneda, Orden.Fecha"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Clave = " + "'" + ZZClave + "'"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            ZVector(Ciclo, 3) = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            ZVector(Ciclo, 4) = IIf(IsNull(rstOrden!Carpeta), "0", rstOrden!Carpeta)
            ZVector(Ciclo, 5) = Str$(rstOrden!Precio)
            ZVector(Ciclo, 6) = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            ZVector(Ciclo, 7) = rstOrden!Fecha
            
            rstOrden.Close
            
                Else
                
            ZVector(Ciclo, 1) = ""
            ZVector(Ciclo, 2) = ""
            ZVector(Ciclo, 8) = ""
        
        End If
            
    Next Ciclo
    
    
    
    Rem ver desde aca
    
    Barra.Min = 1
    Barra.Max = ZLugar
    
    For Ciclo = 1 To ZLugar
        
        Contador.Text = Ciclo
        Barra.Value = Ciclo
        DoEvents
    
        ZZLaudo = ZVector(Ciclo, 1)
        ZZOrden = ZVector(Ciclo, 2)
        ZZTipoOrden = ZVector(Ciclo, 3)
        ZZCarpeta = ZVector(Ciclo, 4)
        ZZPrecio = ZVector(Ciclo, 5)
        ZZMoneda = ZVector(Ciclo, 6)
        ZZFecha = ZVector(Ciclo, 7)
        ZZArticulo = ZVector(Ciclo, 8)
        
        If Val(ZZOrden) <> 0 And Val(ZZLaudo) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select CostoPartida.Laudo"
            ZSql = ZSql + " FROM CostoPartida"
            ZSql = ZSql + " Where CostoPartida.Laudo = " + "'" + ZZLaudo + "'"
            spCostoPartida = ZSql
            Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
            If rstCostoPartida.RecordCount > 0 Then
                rstCostoPartida.Close
                    Else
                If Val(ZZTipoOrden) = 1 Then
                    Call Calcula_Importacion
                        Else
                    If Val(ZZTipoOrden) = 0 Then
                        ZZParidad = 0
                        If Val(ZZMoneda) = 1 Then
                            spCambios = "ConsultaCambio " + "'" + ZZFecha + "'"
                            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCambios.RecordCount > 0 Then
                                ZZParidad = rstCambios!Cambio
                                rstCambios.Close
                            End If
                        End If
                        Call Graba_Orden
                    End If
                End If
            End If
                
        End If
        
    Next Ciclo
    
    Call Cancela_click

End Sub




Private Sub ProcesoAnterior_Click()

    ZLugar = 0
    Erase ZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Laudo"
    ZSql = ZSql + " Order by Laudo.Clave"
    spLaudo = ZSql
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstLaudo!Orden <> 0 And rstLaudo!Laudo <> 0 Then
                    
                        ZZEntra = "S"
                        For ZZCicla = 1 To ZLugar
                            If rstLaudo!Orden = Val(ZVector(ZZCicla, 2)) Then
                                ZZEntra = "N"
                                Exit For
                            End If
                        Next ZZCicla
                        
                        If ZZEntra = "S" Then
                            ZLugar = ZLugar + 1
                            
                            ZVector(ZLugar, 1) = Str$(rstLaudo!Laudo)
                            ZVector(ZLugar, 2) = Str$(rstLaudo!Orden)
                            ZVector(ZLugar, 3) = ""
                            ZVector(ZLugar, 4) = ""
                            ZVector(ZLugar, 5) = ""
                            ZVector(ZLugar, 6) = ""
                            ZVector(ZLugar, 7) = ""
                        End If
                        
                    End If
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZZLaudo = ZVector(Ciclo, 1)
        ZZOrden = ZVector(Ciclo, 2)
        
        Call Ceros(ZZOrden, 6)
        ZZClave = ZZOrden + "01"
    
        ZSql = ""
        ZSql = ZSql + "Select Orden.Orden, Orden.Tipo, Orden.Carpeta, Orden.Precio, Orden.Moneda, Orden.Fecha"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Clave = " + "'" + ZZClave + "'"
        ZSql = ZSql + " Order by Orden.Clave"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
    
            ZVector(Ciclo, 3) = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            ZVector(Ciclo, 4) = IIf(IsNull(rstOrden!Carpeta), "0", rstOrden!Carpeta)
            ZVector(Ciclo, 5) = Str$(rstOrden!Precio)
            ZVector(Ciclo, 6) = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
            ZVector(Ciclo, 7) = rstOrden!Fecha
            
            rstOrden.Close
            
                Else
                
            ZVector(Ciclo, 1) = ""
            ZVector(Ciclo, 2) = ""
        
        End If
            
    Next Ciclo
    
    
    
    Rem ver desde aca
    
    Barra.Min = 1
    Barra.Max = ZLugar
    
    For Ciclo = 1 To ZLugar
    
        Contador.Text = Ciclo
        Barra.Value = Ciclo
        DoEvents
    
        ZZLaudo = ZVector(Ciclo, 1)
        ZZOrden = ZVector(Ciclo, 2)
        ZZTipoOrden = ZVector(Ciclo, 3)
        ZZCarpeta = ZVector(Ciclo, 4)
        ZZPrecio = ZVector(Ciclo, 5)
        ZZMoneda = ZVector(Ciclo, 6)
        ZZFecha = ZVector(Ciclo, 7)
        
        If Val(ZZOrden) <> 0 And Val(ZZLaudo) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select CostoPartida.Laudo"
            ZSql = ZSql + " FROM CostoPartida"
            ZSql = ZSql + " Where CostoPartida.Laudo = " + "'" + ZZLaudo + "'"
            spCostoPartida = ZSql
            Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
            If rstCostoPartida.RecordCount > 0 Then
                rstCostoPartida.Close
                    Else
                If Val(ZZTipoOrden) = 1 Then
                    Call Calcula_Importacion
                        Else
                    If Val(ZZTipoOrden) = 0 Then
                        ZZParidad = 0
                        If Val(ZZMoneda) = 1 Then
                            spCambios = "ConsultaCambio " + "'" + ZZFecha + "'"
                            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCambios.RecordCount > 0 Then
                                ZZParidad = rstCambios!Cambio
                            End If
                        End If
                        Call Graba_Orden
                    End If
                End If
            End If
                
        End If
        
    Next Ciclo
    
    Call Cancela_click

End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub




